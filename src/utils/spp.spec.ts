
import assert from 'assert';
import sinon from 'sinon';
import { spp } from './spp.js';
import { sinonUtil } from './sinonUtil.js';
import request from '../request.js';
import { Logger } from '../cli/Logger.js';

describe('utils/spp', () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const siteUrl = 'https://contoso.sharepoint.com';
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
  });

  it('calls api correctly and throw an error when site is not a content center using assertSiteIsContentCenter', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web?$select=WebTemplateConfiguration`) {
        return {
          WebTemplateConfiguration: 'SITEPAGEPUBLISHING#0'
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(spp.assertSiteIsContentCenter(siteUrl, logger, false), Error('https://contoso.sharepoint.com is not a content site.'));
  });

  it('calls api correctly and does not throw an error when site is a content center using assertSiteIsContentCenter', async () => {
    const stubGet = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web?$select=WebTemplateConfiguration`) {
        return {
          WebTemplateConfiguration: 'CONTENTCTR#0'
        };
      }

      throw 'Invalid request';
    });

    await spp.assertSiteIsContentCenter(siteUrl, logger, false);
    assert(stubGet.calledOnce);
  });

  it('calls api correctly and shows verbose message using assertSiteIsContentCenter', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web?$select=WebTemplateConfiguration`) {
        return {
          WebTemplateConfiguration: 'CONTENTCTR#0'
        };
      }

      throw 'Invalid request';
    });

    await spp.assertSiteIsContentCenter(siteUrl, logger, true);
    assert(loggerLogSpy.calledOnce);
  });


  it('retrieves model by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return model;
      }

      throw 'Invalid request';
    });

    const actual = await spp.getModelById('https://contoso.sharepoint.com/sites/portal', '9b1b1e42-794b-4c71-93ac-5ed92488b67f', logger, false);
    assert.deepStrictEqual(actual, model);
  });

  it('retrieves model by id with verbose message', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        return model;
      }

      throw 'Invalid request';
    });

    await spp.getModelById('https://contoso.sharepoint.com/sites/portal', '9b1b1e42-794b-4c71-93ac-5ed92488b67f', logger, true);
    assert(loggerLogSpy.calledOnce);
  });

  it('retrieves model by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbytitle('modelname.classifier')`) {
        return model;
      }

      throw 'Invalid request';
    });

    const actual = await spp.getModelByTitle('https://contoso.sharepoint.com/sites/portal', 'ModelName', logger, false);
    assert.deepStrictEqual(actual, model);
  });

  it('retrieves model by title with verbose message', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbytitle('modelname.classifier')`) {
        return model;
      }

      throw 'Invalid request';
    });

    await spp.getModelByTitle('https://contoso.sharepoint.com/sites/portal', 'ModelName', logger, true);
    assert(loggerLogSpy.calledOnce);
  });

  it('retrieves model without additional information by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbytitle('modelname.classifier')`) {
        return modelWithoutAdditionalData;
      }

      throw 'Invalid request';
    });

    const actual = await spp.getModelByTitle('https://contoso.sharepoint.com/sites/portal', 'ModelName', logger, false);
    assert.deepStrictEqual(actual, modelWithoutAdditionalData);
  });

  it('retrieves model by title with classifier suffix', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbytitle('modelname.classifier')`) {
        return model;
      }

      throw 'Invalid request';
    });


    const actual = await spp.getModelByTitle('https://contoso.sharepoint.com/sites/portal', 'ModelName.classifier', logger, false);
    assert.deepStrictEqual(actual, model);
  });

  it('correctly handles a model is not found error by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models/getbyuniqueid('9b1b1e42-794b-4c71-93ac-5ed92488b67f')`) {
        throw Error('File Not Found.');
      }

      throw 'Invalid request';
    });

    await assert.rejects(spp.getModelById('https://contoso.sharepoint.com/sites/portal', '9b1b1e42-794b-4c71-93ac-5ed92488b67f', logger, false),
      Error('File Not Found.'));
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

    await assert.rejects(spp.getModelByTitle('https://contoso.sharepoint.com/sites/portal', 'invalidTitle', logger, false),
      Error(`Model 'invalidTitle' was not found.`));
  });
});