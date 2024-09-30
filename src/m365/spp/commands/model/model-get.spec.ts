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
import command from './model-get.js';
import { spp } from '../../../../utils/spp.js';

describe(commands.MODEL_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const model = {
    "AIBuilderHybridModelType": null,
    "AzureCognitivePrebuiltModelName": null,
    "BaseContentTypeName": null,
    "ConfidenceScore": "{\"trainingStatus\":{\"kind\":\"original\",\"ClassifierStatus\":{\"TrainingStatus\":\"success\",\"TimeStamp\":1716547860981},\"ExtractorsStatus\":[{\"TimeStamp\":1716547860173,\"ExtractorName\":\"Name\",\"TrainingStatus\":\"success\"}]},\"modelAccuracy\":{\"Classifier\":1,\"Extractors\":{\"Name\":0.333333343}},\"perSampleAccuracy\":{\"1\":{\"Classifier\":1,\"Extractors\":{\"Name\":1}},\"2\":{\"Classifier\":1,\"Extractors\":{\"Name\":0}},\"3\":{\"Classifier\":1,\"Extractors\":{\"Name\":0}},\"4\":{\"Classifier\":1,\"Extractors\":{\"Name\":0}},\"5\":{\"Classifier\":1,\"Extractors\":{\"Name\":0}},\"6\":{\"Classifier\":1,\"Extractors\":{\"Name\":1}}},\"perSamplePrediction\":{\"1\":{\"Extractors\":{\"Name\":[\"FirstName\"]}},\"2\":{\"Extractors\":{\"Name\":[]}},\"3\":{\"Extractors\":{\"Name\":[]}},\"4\":{\"Extractors\":{\"Name\":[]}},\"5\":{\"Extractors\":{\"Name\":[]}},\"6\":{\"Extractors\":{\"Name\":[]}}},\"trainingFailures\":{}}",
    "ContentTypeGroup": "Intelligent Document Content Types",
    "ContentTypeId": "0x010100A5C3671D1FB1A64D9F280C628D041692",
    "ContentTypeName": "TeachingModel",
    "Created": "2024-05-23T16:51:29Z",
    "CreatedBy": "i:0#.f|membership|user@contoso.onmicrosoft.com",
    "DriveId": "b!qTVLltt1P02LUgejOy6O_1amoFeu1EJBlawH83UtYbQs_H3KVKAcQpuQOpNLl646",
    "Explanations": "{\"Classifier\":[{\"id\":\"950f39cd-5e72-4442-ae8a-ea4bae32962c\",\"kind\":\"regexFeature\",\"name\":\"Email address\",\"active\":true,\"pattern\":\"[A-Za-z0-9._%-]+@[A-Za-z0-9.-]+.[A-Za-z]{2,6}\"},{\"id\":\"d3f2940d-1df1-4ba8-975a-db3a4d626d5c\",\"kind\":\"dictionaryFeature\",\"name\":\"FirstName\",\"active\":true,\"nGrams\":[\"FirstName\"],\"caseSensitive\":false,\"ignoreDigitIdentity\":false,\"ignoreLetterIdentity\":false},{\"id\":\"077966e1-73be-44c9-855f-a4eade6a280b\",\"kind\":\"modelFeature\",\"name\":\"Name\",\"active\":true,\"modelReference\":\"Name\",\"conceptId\":\"309b64e5-acd5-4538-a5b6-c6bfcdc1ffbf\"}],\"Extractors\":{\"Name\":[{\"id\":\"69e412bc-e5bc-4657-b378-34a01966bb92\",\"kind\":\"dictionaryFeature\",\"name\":\"Before label\",\"active\":true,\"nGrams\":[\"Test\",\"'Surname\",\"Test (\"],\"caseSensitive\":false,\"ignoreDigitIdentity\":false,\"ignoreLetterIdentity\":false}]}}",
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
  };

  const modelWithoutAdditionalData = {
    "AIBuilderHybridModelType": null,
    "AzureCognitivePrebuiltModelName": null,
    "BaseContentTypeName": null,
    "ConfidenceScore": null,
    "ContentTypeGroup": "Intelligent Document Content Types",
    "ContentTypeId": "0x010100A5C3671D1FB1A64D9F280C628D041692",
    "ContentTypeName": "TeachingModel",
    "Created": "2024-05-23T16:51:29Z",
    "CreatedBy": "i:0#.f|membership|user@contoso.onmicrosoft.com",
    "DriveId": "b!qTVLltt1P02LUgejOy6O_1amoFeu1EJBlawH83UtYbQs_H3KVKAcQpuQOpNLl646",
    "Explanations": null,
    "ID": 1,
    "LastTrained": "2024-05-24T17:51:16Z",
    "ListID": "ca7dfc2c-a054-421c-9b90-3a934b97ae3a",
    "ModelSettings": null,
    "ModelType": 2,
    "Modified": "2024-05-24T17:51:02Z",
    "ModifiedBy": "i:0#.f|membership|user@contoso.onmicrosoft.com",
    "ObjectId": "01HQDCWVGIEBDRN3RVK5A3UJW3M4TCMT45",
    "PublicationType": 0,
    "Schemas": null,
    "SourceSiteUrl": "https://contoso.sharepoint.com/sites/SyntexTest",
    "SourceUrl": null,
    "SourceWebServerRelativeUrl": "/sites/SyntexTest",
    "UniqueId": "164720c8-35ee-4157-ba26-db6726264f9d"
  };

  const modelResult = {
    "AIBuilderHybridModelType": null,
    "AzureCognitivePrebuiltModelName": null,
    "BaseContentTypeName": null,
    "ConfidenceScore": { "trainingStatus": { "kind": "original", "ClassifierStatus": { "TrainingStatus": "success", "TimeStamp": 1716547860981 }, "ExtractorsStatus": [{ "TimeStamp": 1716547860173, "ExtractorName": "Name", "TrainingStatus": "success" }] }, "modelAccuracy": { "Classifier": 1, "Extractors": { "Name": 0.333333343 } }, "perSampleAccuracy": { "1": { "Classifier": 1, "Extractors": { "Name": 1 } }, "2": { "Classifier": 1, "Extractors": { "Name": 0 } }, "3": { "Classifier": 1, "Extractors": { "Name": 0 } }, "4": { "Classifier": 1, "Extractors": { "Name": 0 } }, "5": { "Classifier": 1, "Extractors": { "Name": 0 } }, "6": { "Classifier": 1, "Extractors": { "Name": 1 } } }, "perSamplePrediction": { "1": { "Extractors": { "Name": ["FirstName"] } }, "2": { "Extractors": { "Name": [] } }, "3": { "Extractors": { "Name": [] } }, "4": { "Extractors": { "Name": [] } }, "5": { "Extractors": { "Name": [] } }, "6": { "Extractors": { "Name": [] } } }, "trainingFailures": {} },
    "ContentTypeGroup": "Intelligent Document Content Types",
    "ContentTypeId": "0x010100A5C3671D1FB1A64D9F280C628D041692",
    "ContentTypeName": "TeachingModel",
    "Created": "2024-05-23T16:51:29Z",
    "CreatedBy": "i:0#.f|membership|user@contoso.onmicrosoft.com",
    "DriveId": "b!qTVLltt1P02LUgejOy6O_1amoFeu1EJBlawH83UtYbQs_H3KVKAcQpuQOpNLl646",
    "Explanations": { "Classifier": [{ "id": "950f39cd-5e72-4442-ae8a-ea4bae32962c", "kind": "regexFeature", "name": "Email address", "active": true, "pattern": "[A-Za-z0-9._%-]+@[A-Za-z0-9.-]+.[A-Za-z]{2,6}" }, { "id": "d3f2940d-1df1-4ba8-975a-db3a4d626d5c", "kind": "dictionaryFeature", "name": "FirstName", "active": true, "nGrams": ["FirstName"], "caseSensitive": false, "ignoreDigitIdentity": false, "ignoreLetterIdentity": false }, { "id": "077966e1-73be-44c9-855f-a4eade6a280b", "kind": "modelFeature", "name": "Name", "active": true, "modelReference": "Name", "conceptId": "309b64e5-acd5-4538-a5b6-c6bfcdc1ffbf" }], "Extractors": { "Name": [{ "id": "69e412bc-e5bc-4657-b378-34a01966bb92", "kind": "dictionaryFeature", "name": "Before label", "active": true, "nGrams": ["Test", "'Surname", "Test ("], "caseSensitive": false, "ignoreDigitIdentity": false, "ignoreLetterIdentity": false }] } },
    "ID": 1,
    "LastTrained": "2024-05-24T17:51:16Z",
    "ListID": "ca7dfc2c-a054-421c-9b90-3a934b97ae3a",
    "ModelSettings": { "ModelOrigin": { "LibraryId": "8a6027ab-c584-4394-ba9c-3dc4dd152b65", "Published": false, "SelecteFileUniqueIds": [] } },
    "ModelType": 2,
    "Modified": "2024-05-24T17:51:02Z",
    "ModifiedBy": "i:0#.f|membership|user@contoso.onmicrosoft.com",
    "ObjectId": "01HQDCWVGIEBDRN3RVK5A3UJW3M4TCMT45",
    "PublicationType": 0,
    "Schemas": { "Version": 2, "Extractors": { "Name": { "concepts": { "309b64e5-acd5-4538-a5b6-c6bfcdc1ffbf": { "name": "Name" } }, "relationships": [], "id": "Name" } } },
    "SourceSiteUrl": "https://contoso.sharepoint.com/sites/SyntexTest",
    "SourceUrl": null,
    "SourceWebServerRelativeUrl": "/sites/SyntexTest",
    "UniqueId": "164720c8-35ee-4157-ba26-db6726264f9d"
  };

  const modelResultWithoutAdditionalData = {
    "AIBuilderHybridModelType": null,
    "AzureCognitivePrebuiltModelName": null,
    "BaseContentTypeName": null,
    "ConfidenceScore": null,
    "ContentTypeGroup": "Intelligent Document Content Types",
    "ContentTypeId": "0x010100A5C3671D1FB1A64D9F280C628D041692",
    "ContentTypeName": "TeachingModel",
    "Created": "2024-05-23T16:51:29Z",
    "CreatedBy": "i:0#.f|membership|user@contoso.onmicrosoft.com",
    "DriveId": "b!qTVLltt1P02LUgejOy6O_1amoFeu1EJBlawH83UtYbQs_H3KVKAcQpuQOpNLl646",
    "Explanations": null,
    "ID": 1,
    "LastTrained": "2024-05-24T17:51:16Z",
    "ListID": "ca7dfc2c-a054-421c-9b90-3a934b97ae3a",
    "ModelSettings": null,
    "ModelType": 2,
    "Modified": "2024-05-24T17:51:02Z",
    "ModifiedBy": "i:0#.f|membership|user@contoso.onmicrosoft.com",
    "ObjectId": "01HQDCWVGIEBDRN3RVK5A3UJW3M4TCMT45",
    "PublicationType": 0,
    "SourceSiteUrl": "https://contoso.sharepoint.com/sites/SyntexTest",
    "Schemas": null,
    "SourceUrl": null,
    "SourceWebServerRelativeUrl": "/sites/SyntexTest",
    "UniqueId": "164720c8-35ee-4157-ba26-db6726264f9d"
  };

  const publications = [
    {
      "Created": "2020-11-03T12:52:12Z",
      "CreatedBy": "i:0#.f|membership|user@contoso.onmicrosoft.com",
      "DriveId": "b!7w273MPU8kiqSKc6SWaUkQ3wyRgeHbxDh-e6ShP9X-e-6B-bO36LRI87VfSrxwy9",
      "HasTargetSitePermission": true,
      "ID": 1,
      "ModelId": 3,
      "ModelName": "BenefitsChangeNotice.classifier",
      "ModelType": 0,
      "ModelUniqueId": "b10e0de5-c069-46f9-90f7-4fb8ac001372",
      "ModelVersion": "2.0",
      "Modified": "2020-11-03T18:45:18Z",
      "ModifiedBy": "i:0#.f|membership|user@contoso.onmicrosoft.com",
      "ObjectId": "01XHOXSCVXVZCQAAQWCNGK6NGADKKZUMW7",
      "PublicationType": 1,
      "TargetLibraryId": "7eb6c306-1680-4ba9-9bbe-6b5b7efb27be",
      "TargetLibraryName": "Documents",
      "TargetLibraryRemoved": false,
      "TargetLibraryServerRelativeUrl": "/sites/SyntexTest/Shared%20Documents",
      "TargetLibraryUrl": "https://contoso.sharepoint.com/sites/SyntexTest/Shared%20Documents",
      "TargetSiteId": "7e44de0a-30e3-45ab-a601-057298c9068f",
      "TargetSiteUrl": "https://contoso.sharepoint.com/sites/SyntexTest",
      "TargetTableListId": "00000000-0000-0000-0000-000000000000",
      "TargetTableListName": null,
      "TargetTableListRemoved": false,
      "TargetTableListServerRelativeUrl": null,
      "TargetTableListUrl": null,
      "TargetWebId": "9166df43-8e45-43b8-94ef-3a22826346de",
      "TargetWebName": "SyntexTest",
      "TargetWebServerRelativeUrl": "/sites/SyntexTest",
      "UniqueId": "0045aeb7-1602-4c13-af34-c01a959a32df",
      "ViewOption": "NewViewAsDefault"
    }
  ];

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spp, 'assertSiteIsContentCenter').resolves();
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
    assert.strictEqual(command.name, commands.MODEL_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when required parameters are valid with id', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when required parameters are valid with title', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', title: 'ModelName' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when required parameters are valid with id and withPublications', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', withPublications: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when required parameters are valid with title and withPublications', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', title: 'ModelName', withPublications: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when siteUrl is not valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'invalidUrl', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when id is not valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales', id: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('correctly handles a model is not found error by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        throw {
          error: {
            "odata.error": {
              code: "-1, Microsoft.Office.Server.ContentCenter.ModelNotFoundException",
              message: {
                lang: "en-US",
                value: "File Not Found."
              }
            }
          }
        };
      }
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }),
      new CommandError('File Not Found.'));
  });

  it('correctly handles a model is not found error by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbytitle('invalidtitle.classifier')`) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, siteUrl: 'https://contoso.sharepoint.com/sites/portal', title: 'invalidTitle' } }),
      new CommandError('Model not found.'));
  });

  it('retrieves model by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return model;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], modelResult);
  });

  it('retrieves model by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbytitle('modelname.classifier')`) {
        return model;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal', title: 'ModelName' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], modelResult);
  });

  it('retrieves model without additional information by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbytitle('modelname.classifier')`) {
        return modelWithoutAdditionalData;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal', title: 'ModelName' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], modelResultWithoutAdditionalData);
  });

  it('retrieves model by title with classifier suffix', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbytitle('modelname.classifier')`) {
        return model;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal', title: 'ModelName.classifier' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], modelResult);
  });

  it('gets correct model when the site URL has a trailing slash', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return model;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal/', id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], modelResult);
  });

  it('retrieves model by id with withPublications', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('164720c8-35ee-4157-ba26-db6726264f9d')`) {
        return model;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/publications/getbymodeluniqueid('164720c8-35ee-4157-ba26-db6726264f9d')`) {
        return { value: publications };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal', id: '164720c8-35ee-4157-ba26-db6726264f9d', withPublications: true } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], { ...modelResult, Publications: publications });
  });

  it('retrieves model by title with withPublications', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbytitle('modelname.classifier')`) {
        return model;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/publications/getbymodeluniqueid('164720c8-35ee-4157-ba26-db6726264f9d')`) {
        return { value: publications };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal', title: 'ModelName', withPublications: true } });
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], { ...modelResult, Publications: publications });
  });
});