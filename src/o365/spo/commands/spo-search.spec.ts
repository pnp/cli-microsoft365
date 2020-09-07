import commands from '../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth from '../../../Auth';
const command: Command = require('./spo-search');
import * as assert from 'assert';
import request from '../../../request';
import Utils from '../../../Utils';
import { ResultTableRow } from './search/datatypes/ResultTableRow';
import { SearchResult } from './search/datatypes/SearchResult';

enum TestID {
  None,
  QueryAll_NoParameterTest,
  QueryAll_WithQueryTemplateTest,
  QueryDocuments_WithStartRow0Test,
  QueryDocuments_WithStartRow1Test,
  QueryDocuments_NoStartRowTest,
  QueryDocuments_NoParameterTest,
  QueryAll_WithRowLimitTest,
  QueryAll_WithSourceIdTest,
  QueryAll_WithTrimDuplicatesTest,
  QueryAll_WithEnableStemmingTest,
  QueryAll_WithCultureTest,
  QueryAll_WithRefinementFiltersTest,
  QueryAll_SortListTest,
  QueryAll_WithRankingModelIdTest,
  QueryAll_WithStartRowTest,
  QueryAll_WithPropertiesTest,
  QueryAll_WithSourceNameAndPreviousPropertiesTest,
  QueryAll_WithSourceNameAndNoPreviousPropertiesTest,
  QueryAll_WithRefinersTest,
  QueryAll_WithWebTest,
  QueryAll_WithHiddenConstraintsTest,
  QueryAll_WithClientTypeTest,
  QueryAll_WithEnablePhoneticTest,
  QueryAll_WithProcessBestBetsTest,
  QueryAll_WithEnableQueryRulesTest,
  QueryAll_WithProcessPersonalFavoritesTest
}

describe(commands.SEARCH, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let returnArrayLength = 0;
  let executedTest: TestID = TestID.None;
  let urlContains = (opts: any, substring: string): boolean => {
    return opts.url.toUpperCase().indexOf(substring.toUpperCase()) > -1;
  }
  let filterRows = (rows: ResultTableRow[], key: string, value: string) => {
    return rows.filter(row => {
      return row.Cells.filter(cell => {
        return (cell.Key.toUpperCase() === key.toUpperCase() && cell.Value.toUpperCase() === value.toUpperCase());
      }).length > 0;
    });
  }
  let getFakeRows = (): ResultTableRow[] => {
    return [
      {
        "Cells": [
          { "Key": "Rank", "Value": "1", "ValueType": "Edm.Double" },
          { "Key": "DocId", "Value": "1", "ValueType": "Edm.Int64" },
          { "Key": "Path", "Value": "MyPath-item1", "ValueType": "Edm.String" },
          { "Key": "Author", "Value": "myAuthor-item1", "ValueType": "Edm.String" },
          { "Key": "FileType", "Value": "docx", "ValueType": "Edm.String" },
          { "Key": "OriginalPath", "Value": "myOriginalPath-item1", "ValueType": "Edm.String" },
          { "Key": "PartitionId", "Value": "00000000-0000-0000-0000-000000000000", "ValueType": "Edm.Guid" },
          { "Key": "UrlZone", "Value": "0", "ValueType": "Edm.Int32" },
          { "Key": "Culture", "Value": "en-US", "ValueType": "Edm.String" },
          { "Key": "ResultTypeId", "Value": "0", "ValueType": "Edm.Int32" },
          { "Key": "IsDocument", "Value": "true", "ValueType": "Edm.Boolean" },
          { "Key": "RenderTemplateId", "Value": "~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_Default.js", "ValueType": "Edm.String" }
        ]
      },
      {
        "Cells": [
          { "Key": "Rank", "Value": "2", "ValueType": "Edm.Double" },
          { "Key": "DocId", "Value": "2", "ValueType": "Edm.Int64" },
          { "Key": "Path", "Value": "MyPath-item2", "ValueType": "Edm.String" },
          { "Key": "Author", "Value": "myAuthor-item2", "ValueType": "Edm.String" },
          { "Key": "FileType", "Value": "docx", "ValueType": "Edm.String" },
          { "Key": "OriginalPath", "Value": "myOriginalPath-item2", "ValueType": "Edm.String" },
          { "Key": "PartitionId", "Value": "00000000-0000-0000-0000-000000000000", "ValueType": "Edm.Guid" },
          { "Key": "UrlZone", "Value": "0", "ValueType": "Edm.Int32" },
          { "Key": "Culture", "Value": "en-US", "ValueType": "Edm.String" },
          { "Key": "ResultTypeId", "Value": "0", "ValueType": "Edm.Int32" },
          { "Key": "IsDocument", "Value": "true", "ValueType": "Edm.Boolean" },
          { "Key": "RenderTemplateId", "Value": "~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_Default.js", "ValueType": "Edm.String" }
        ]
      },
      {
        "Cells": [
          { "Key": "Rank", "Value": "3", "ValueType": "Edm.Double" },
          { "Key": "DocId", "Value": "3", "ValueType": "Edm.Int64" },
          { "Key": "Path", "Value": "MyPath-item3", "ValueType": "Edm.String" },
          { "Key": "Author", "Value": "myAuthor-item3", "ValueType": "Edm.String" },
          { "Key": "FileType", "Value": "aspx", "ValueType": "Edm.String" },
          { "Key": "OriginalPath", "Value": "myOriginalPath-item3", "ValueType": "Edm.String" },
          { "Key": "PartitionId", "Value": "00000000-0000-0000-0000-000000000000", "ValueType": "Edm.Guid" },
          { "Key": "UrlZone", "Value": "0", "ValueType": "Edm.Int32" },
          { "Key": "Culture", "Value": "en-US", "ValueType": "Edm.String" },
          { "Key": "ResultTypeId", "Value": "0", "ValueType": "Edm.Int32" },
          { "Key": "IsDocument", "Value": "false", "ValueType": "Edm.Boolean" },
          { "Key": "RenderTemplateId", "Value": "~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_Default.js", "ValueType": "Edm.String" }
        ]
      },
      {
        "Cells": [
          { "Key": "Rank", "Value": "4", "ValueType": "Edm.Double" },
          { "Key": "DocId", "Value": "4", "ValueType": "Edm.Int64" },
          { "Key": "Path", "Value": "MyPath-item4", "ValueType": "Edm.String" },
          { "Key": "Author", "Value": "myAuthor-item4", "ValueType": "Edm.String" },
          { "Key": "FileType", "Value": "aspx", "ValueType": "Edm.String" },
          { "Key": "OriginalPath", "Value": "myOriginalPath-item4", "ValueType": "Edm.String" },
          { "Key": "PartitionId", "Value": "00000000-0000-0000-0000-000000000000", "ValueType": "Edm.Guid" },
          { "Key": "UrlZone", "Value": "0", "ValueType": "Edm.Int32" },
          { "Key": "Culture", "Value": "nl-NL", "ValueType": "Edm.String" },
          { "Key": "ResultTypeId", "Value": "0", "ValueType": "Edm.Int32" },
          { "Key": "IsDocument", "Value": "false", "ValueType": "Edm.Boolean" },
          { "Key": "RenderTemplateId", "Value": "~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_Default.js", "ValueType": "Edm.String" }
        ]
      }
    ];
  };
  let fakeRows: ResultTableRow[] = getFakeRows();
  let getQueryResult = (rows: ResultTableRow[], totalRows?: number): SearchResult => {
    returnArrayLength = totalRows ? totalRows : rows.length;

    return {
      "ElapsedTime": 83,
      "PrimaryQueryResult": {
        "CustomResults": [],
        "QueryId": "00000000-0000-0000-0000-000000000000",
        "QueryRuleId": "00000000-0000-0000-0000-000000000000",
        "RefinementResults": null,
        "RelevantResults": {
          "GroupTemplateId": null,
          "ItemTemplateId": null,
          "Properties": [
            {
              "Key": "GenerationId",
              "Value": "9223372036854775806",
              "ValueType": "Edm.Int64"
            }
          ],
          "ResultTitle": null,
          "ResultTitleUrl": null,
          "RowCount": rows.length,
          "Table": {
            "Rows": fakeRows
          },
          "TotalRows": returnArrayLength,
          "TotalRowsIncludingDuplicates": returnArrayLength
        },
        "SpecialTermResults": null
      },
      "Properties": [
        {
          "Key": "RowLimit",
          "Value": "10",
          "ValueType": "Edm.Int32"
        }
      ],
      "SecondaryQueryResults": [],
      "SpellingSuggestion": "",
      "TriggeredRules": []
    };
  }
  let getFakes = (opts: any) => {
    if (urlContains(opts, 'QUERYTEXT=\'ISDOCUMENT:1\'')) {
      let rows = filterRows(fakeRows, 'ISDOCUMENT', 'TRUE');

      if (urlContains(opts, 'ROWLIMIT=1')) {
        if (urlContains(opts, 'STARTROW=0')) {
          executedTest = TestID.QueryDocuments_WithStartRow0Test;
          return Promise.resolve(getQueryResult([rows[0]], 2));
        }
        else if (urlContains(opts, 'STARTROW=1')) {
          executedTest = TestID.QueryDocuments_WithStartRow1Test;
          return Promise.resolve(getQueryResult([rows[1]], 2));
        }
        else {
          executedTest = TestID.QueryDocuments_NoStartRowTest;
          return Promise.resolve(getQueryResult([]));
        }
      }

      executedTest = TestID.QueryDocuments_NoParameterTest;
      return Promise.resolve(getQueryResult(rows));
    }
    if (urlContains(opts, 'QUERYTEXT=\'*\'')) {
      let rows = fakeRows;
      if (urlContains(opts, 'ROWLIMIT=1')) {
        executedTest = TestID.QueryAll_WithRowLimitTest;
        return Promise.resolve(getQueryResult([rows[0]]));
      }
      if (urlContains(opts, 'SOURCEID=\'6E71030E-5E16-4406-9BFF-9C1829843083\'')) {
        executedTest = TestID.QueryAll_WithSourceIdTest;
        return Promise.resolve(getQueryResult([rows[3]]));
      }
      if (urlContains(opts, 'TRIMDUPLICATES=TRUE')) {
        executedTest = TestID.QueryAll_WithTrimDuplicatesTest;
        return Promise.resolve(getQueryResult([rows[2], rows[3]]));
      }
      if (urlContains(opts, 'ENABLESTEMMING=FALSE')) {
        executedTest = TestID.QueryAll_WithEnableStemmingTest;
        return Promise.resolve(getQueryResult([rows[2], rows[3]]));
      }
      if (urlContains(opts, 'CULTURE=1043')) {
        rows = filterRows(fakeRows, 'CULTURE', 'NL-NL');

        executedTest = TestID.QueryAll_WithCultureTest;
        return Promise.resolve(getQueryResult(rows));
      }
      if (urlContains(opts, 'refinementfilters=\'fileExtension:equals("docx")\'')) {
        rows = filterRows(fakeRows, 'FILETYPE', 'DOCX');

        executedTest = TestID.QueryAll_WithRefinementFiltersTest;
        return Promise.resolve(getQueryResult(rows));
      }
      if (urlContains(opts, 'queryTemplate=\'{searchterms} fileType:docx\'')) {
        rows = filterRows(fakeRows, 'FILETYPE', 'DOCX');

        executedTest = TestID.QueryAll_WithQueryTemplateTest;
        return Promise.resolve(getQueryResult(rows));
      }
      if (urlContains(opts, 'sortList=\'Rank%3Aascending\'')) {
        executedTest = TestID.QueryAll_SortListTest;
        return Promise.resolve(getQueryResult(fakeRows));
      }
      if (urlContains(opts, 'rankingModelId=\'d4ac6500-d1d0-48aa-86d4-8fe9a57a74af\'')) {
        executedTest = TestID.QueryAll_WithRankingModelIdTest;
        return Promise.resolve(getQueryResult(fakeRows));
      }
      if (urlContains(opts, 'startRow=1')) {
        executedTest = TestID.QueryAll_WithStartRowTest;
        var rowsToReturn = fakeRows.slice();
        rowsToReturn.splice(0, 1);
        return Promise.resolve(getQueryResult(rowsToReturn));
      }
      if (urlContains(opts, 'properties=\'termid:guid\'')) {
        executedTest = TestID.QueryAll_WithPropertiesTest;
        return Promise.resolve(getQueryResult(fakeRows));
      }
      if (urlContains(opts, 'properties=\'SourceName:Local SharePoint Results,SourceLevel:SPSite\'')) {
        executedTest = TestID.QueryAll_WithSourceNameAndNoPreviousPropertiesTest;
        return Promise.resolve(getQueryResult(fakeRows));
      }
      if (urlContains(opts, 'properties=\'some:property,SourceName:Local SharePoint Results,SourceLevel:SPSite\'')) {
        executedTest = TestID.QueryAll_WithSourceNameAndPreviousPropertiesTest;
        return Promise.resolve(getQueryResult(fakeRows));
      }
      if (urlContains(opts, 'refiners=\'author,size\'')) {
        executedTest = TestID.QueryAll_WithRefinersTest;
        return Promise.resolve(getQueryResult(fakeRows));
      }
      if (urlContains(opts, 'https://contoso.sharepoint.com/sites/subsite')) {
        executedTest = TestID.QueryAll_WithWebTest;
        return Promise.resolve(getQueryResult(fakeRows));
      }
      if (urlContains(opts, 'hiddenConstraints=\'developer\'')) {
        executedTest = TestID.QueryAll_WithHiddenConstraintsTest;
        return Promise.resolve(getQueryResult(fakeRows));
      }
      if (urlContains(opts, 'clientType=\'custom\'')) {
        executedTest = TestID.QueryAll_WithClientTypeTest;
        return Promise.resolve(getQueryResult(fakeRows));
      }

      if (urlContains(opts, 'enablephonetic=true')) {
        executedTest = TestID.QueryAll_WithEnablePhoneticTest;
        return Promise.resolve(getQueryResult(fakeRows));
      }
      if (urlContains(opts, 'processBestBets=true')) {
        executedTest = TestID.QueryAll_WithProcessBestBetsTest;
        return Promise.resolve(getQueryResult(fakeRows));
      }
      if (urlContains(opts, 'enableQueryRules=false')) {
        executedTest = TestID.QueryAll_WithEnableQueryRulesTest;
        return Promise.resolve(getQueryResult(fakeRows));
      }
      if (urlContains(opts, 'processPersonalFavorites=true')) {
        executedTest = TestID.QueryAll_WithProcessPersonalFavoritesTest;
        return Promise.resolve(getQueryResult(fakeRows));
      }

      executedTest = TestID.QueryAll_NoParameterTest;
      return Promise.resolve(getQueryResult(rows));
    }
    returnArrayLength = 0;
    return Promise.reject('Invalid request');
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    vorpal = require('../../../vorpal-init');
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
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.SEARCH), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('executes search request', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'json',
        debug: true,
        query: '*'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_NoParameterTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with output option text', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: 'IsDocument:1'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 2);
        assert.equal(executedTest, TestID.QueryDocuments_NoParameterTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with output option text and \'allResults\'', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: 'IsDocument:1',
        allResults: true,
        rowLimit: 1
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 2);
        assert.equal(executedTest, TestID.QueryDocuments_WithStartRow1Test);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with trimDuplicates', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: '*',
        trimDuplicates: true
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 2);
        assert.equal(executedTest, TestID.QueryAll_WithTrimDuplicatesTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with sortList', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: '*',
        sortList: 'Rank:ascending'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_SortListTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with enableStemming=false', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: '*',
        enableStemming: false
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 2);
        assert.equal(executedTest, TestID.QueryAll_WithEnableStemmingTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with enableStemming=true', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: '*',
        enableStemming: true
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_NoParameterTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with culture', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: '*',
        culture: 1043
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 1);
        assert.equal(executedTest, TestID.QueryAll_WithCultureTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with output option json and \'allResults\'', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'json',
        debug: false,
        query: 'IsDocument:1',
        allResults: true,
        rowLimit: 1
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 2);
        assert.equal(executedTest, TestID.QueryDocuments_WithStartRow1Test);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with selectProperties', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: 'IsDocument:1',
        selectProperties: 'Path'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 2);
        assert.equal(executedTest, TestID.QueryDocuments_NoParameterTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with refinementFilters', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: '*',
        refinementFilters: 'fileExtension:equals("docx")'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 2);
        assert.equal(executedTest, TestID.QueryAll_WithRefinementFiltersTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with queryTemplate', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: '*',
        queryTemplate: '{searchterms} fileType:docx'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 2);
        assert.equal(executedTest, TestID.QueryAll_WithQueryTemplateTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with sourceId', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: '*',
        sourceId: '6e71030e-5e16-4406-9bff-9c1829843083'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 1);
        assert.equal(executedTest, TestID.QueryAll_WithSourceIdTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with rankingModelId', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: '*',
        rankingModelId: 'd4ac6500-d1d0-48aa-86d4-8fe9a57a74af'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_WithRankingModelIdTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with rowLimits defined', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true,
        query: '*',
        rowLimit: 1
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 1);
        assert.equal(executedTest, TestID.QueryAll_WithRowLimitTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with startRow defined', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true,
        query: '*',
        startRow: 1
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 3);
        assert.equal(executedTest, TestID.QueryAll_WithStartRowTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with properties defined', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true,
        query: '*',
        properties: 'termid:guid'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_WithPropertiesTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with sourceName defined and no previous properties', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true,
        query: '*',
        sourceName: 'Local SharePoint Results'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_WithSourceNameAndNoPreviousPropertiesTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with sourceName defined and previous properties (ends with \',\')', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true,
        query: '*',
        properties: 'some:property,',
        sourceName: 'Local SharePoint Results'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_WithSourceNameAndPreviousPropertiesTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with sourceName defined and previous properties (Doesn\'t end with \',\')', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true,
        query: '*',
        properties: 'some:property',
        sourceName: 'Local SharePoint Results'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_WithSourceNameAndPreviousPropertiesTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with refiners defined', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true,
        query: '*',
        refiners: 'author,size'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_WithRefinersTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with web defined', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true,
        query: '*',
        webUrl: 'https://contoso.sharepoint.com/sites/subsite'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_WithWebTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with hiddenConstraints defined', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true,
        query: '*',
        hiddenConstraints: 'developer'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_WithHiddenConstraintsTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with clientType defined', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true,
        query: '*',
        clientType: 'custom'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_WithClientTypeTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with enablePhonetic defined', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true,
        query: '*',
        enablePhonetic: true
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_WithEnablePhoneticTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with processBestBets defined', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true,
        query: '*',
        processBestBets: true
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_WithProcessBestBetsTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with enableQueryRules defined', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true,
        query: '*',
        enableQueryRules: false
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_WithEnableQueryRulesTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with processPersonalFavorites defined', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true,
        query: '*',
        processPersonalFavorites: true
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_WithProcessPersonalFavoritesTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request with parameter rawOutput', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    cmdInstance.action({
      options: {
        output: 'json',
        debug: true,
        query: '*',
        rawOutput: true
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 4);
        assert.equal(executedTest, TestID.QueryAll_NoParameterTest);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the sourceId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        sourceId: '123',
        query: '*'
      }
    });
    assert.notEqual(actual, true);
  });

  it('passes validation if the sourceId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        sourceId: '1caf7dcd-7e83-4c3a-94f7-932a1299c844',
        query: '*'
      }
    });
    assert.equal(actual, true);
  });

  it('fails validation if the rankingModelId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        rankingModelId: '123',
        query: '*'
      }
    });
    assert.notEqual(actual, true);
  });

  it('passes validation if the rankingModelId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        rankingModelId: 'd4ac6500-d1d0-48aa-86d4-8fe9a57a74af',
        query: '*'
      }
    });
    assert.equal(actual, true);
  });

  it('fails validation if the rowLimit is not a valid number', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        rowLimit: '1X',
        query: '*'
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if the startRow is not a valid number', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        startRow: '1X',
        query: '*'
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if the culture is not a valid number', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        culture: '1X',
        query: '*'
      }
    });
    assert.notEqual(actual, true);
  });

  it('command correctly handles reject request', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }

      return Promise.reject('Invalid request');
    });

    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/webs') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
      }
    }, (error?: any) => {
      try {
        assert.equal(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('supports specifying query', () => {
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<query>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the query option is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('passes validation if all options are provided', () => {
    const actual = (command.validate() as CommandValidate)({ options: { query: '*' } });
    assert.equal(actual, true);
  });

  it('fails validation if sortList is in an invalid format', () => {
    const actual = (command.validate() as CommandValidate)({ options: { query: '*', sortList: 'property1:wrongvalue' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if sortList is in a valid format', () => {
    const actual = (command.validate() as CommandValidate)({ options: { query: '*', sortList: 'property1:ascending,property2:descending' } });
    assert.equal(actual, true);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.SEARCH));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
}); 