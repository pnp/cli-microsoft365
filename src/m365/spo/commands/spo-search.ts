import { isNumber } from 'util';
import { cli } from '../../../cli/cli.js';
import { Logger } from '../../../cli/Logger.js';
import GlobalOptions from '../../../GlobalOptions.js';
import request from '../../../request.js';
import { formatting } from '../../../utils/formatting.js';
import { spo } from '../../../utils/spo.js';
import { validation } from '../../../utils/validation.js';
import SpoCommand from '../../base/SpoCommand.js';
import commands from '../commands.js';
import { ResultTableRow } from './search/datatypes/ResultTableRow.js';
import { SearchResult } from './search/datatypes/SearchResult.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  allResults?: boolean;
  clientType?: string;
  culture?: number;
  enablePhonetic?: boolean;
  enableQueryRules?: boolean;
  enableStemming?: boolean;
  hiddenConstraints?: string;
  processBestBets?: boolean;
  processPersonalFavorites?: boolean;
  properties?: string;
  queryText: string;
  queryTemplate?: string;
  rankingModelId?: string;
  rawOutput?: boolean;
  refinementFilters?: string;
  refiners?: string;
  rowLimit?: number;
  selectProperties?: string;
  startRow?: number;
  sortList?: string;
  sourceId?: string;
  sourceName?: string;
  trimDuplicates?: boolean;
  webUrl?: string;
}

class SpoSearchCommand extends SpoCommand {
  public get name(): string {
    return commands.SEARCH;
  }

  public get description(): string {
    return 'Executes a search query';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        selectproperties: typeof args.options.selectProperties !== 'undefined',
        allResults: args.options.allResults,
        rowLimit: args.options.rowLimit,
        sourceId: typeof args.options.sourceId !== 'undefined',
        trimDuplicates: args.options.trimDuplicates,
        enableStemming: args.options.enableStemming,
        culture: args.options.culture,
        refinementFilters: typeof args.options.refinementFilters !== 'undefined',
        queryTemplate: typeof args.options.queryTemplate !== 'undefined',
        sortList: typeof args.options.sortList !== 'undefined',
        rankingModelId: typeof args.options.rankingModelId !== 'undefined',
        startRow: typeof args.options.startRow !== 'undefined',
        properties: typeof args.options.properties !== 'undefined',
        sourceName: typeof args.options.sourceName !== 'undefined',
        refiners: typeof args.options.refiners !== 'undefined',
        webUrl: typeof args.options.webUrl !== 'undefined',
        hiddenConstraints: typeof args.options.hiddenConstraints !== 'undefined',
        clientType: typeof args.options.clientType !== 'undefined',
        enablePhonetic: args.options.enablePhonetic,
        processBestBets: args.options.processBestBets,
        enableQueryRules: args.options.enableQueryRules,
        processPersonalFavorites: args.options.processPersonalFavorites,
        rawOutput: args.options.rawOutput
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-q, --queryText <queryText>'
      },
      {
        option: '-p, --selectProperties [selectProperties]'
      },
      {
        option: '-u, --webUrl [webUrl]'
      },
      {
        option: '--allResults'
      },
      {
        option: '--rowLimit [rowLimit]'
      },
      {
        option: '--sourceId [sourceId]'
      },
      {
        option: '--trimDuplicates'
      },
      {
        option: '--enableStemming'
      },
      {
        option: '--culture [culture]'
      },
      {
        option: '--refinementFilters [refinementFilters]'
      },
      {
        option: '--queryTemplate [queryTemplate]'
      },
      {
        option: '--sortList [sortList]'
      },
      {
        option: '--rankingModelId [rankingModelId]'
      },
      {
        option: '--startRow [startRow]'
      },
      {
        option: '--properties [properties]'
      },
      {
        option: '--sourceName [sourceName]'
      },
      {
        option: '--refiners [refiners]'
      },
      {
        option: '--hiddenConstraints [hiddenConstraints]'
      },
      {
        option: '--clientType [clientType]'
      },
      {
        option: '--enablePhonetic'
      },
      {
        option: '--processBestBets'
      },
      {
        option: '--enableQueryRules'
      },
      {
        option: '--processPersonalFavorites'
      },
      {
        option: '--rawOutput'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.sourceId && !validation.isValidGuid(args.options.sourceId)) {
          return `${args.options.sourceId} is not a valid GUID`;
        }

        if (args.options.rankingModelId && !validation.isValidGuid(args.options.rankingModelId)) {
          return `${args.options.rankingModelId} is not a valid GUID`;
        }

        if (args.options.sortList && !/^([a-z0-9_]+:(ascending|descending))(,([a-z0-9_]+:(ascending|descending)))*$/gi.test(args.options.sortList)) {
          return `sortlist parameter value '${args.options.sortList}' does not match the required pattern (=comma-separated list of '<property>:(ascending|descending)'-pattern)`;
        }

        if (args.options.rowLimit && !isNumber(args.options.rowLimit)) {
          return `${args.options.rowLimit} is not a valid number`;
        }

        if (args.options.startRow && !isNumber(args.options.startRow)) {
          return `${args.options.startRow} is not a valid number`;
        }

        if (args.options.startRow && args.options.allResults) {
          return 'You cannot specify startRow when allResults is set';
        }

        if (args.options.culture && !isNumber(args.options.culture)) {
          return `${args.options.culture} is not a valid number`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let webUrl: string = '';

    try {
      if (args.options.webUrl) {
        webUrl = args.options.webUrl;
      }
      else {
        webUrl = await spo.getSpoUrl(logger, this.debug);
      }

      if (this.verbose) {
        await logger.logToStderr(`Executing search query '${args.options.queryText}' on site at ${webUrl}...`);
      }

      const results: SearchResult[] = await this.executeSearchQuery(logger, args, webUrl, []);
      await this.printResults(logger, args, results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async executeSearchQuery(logger: Logger, args: CommandArgs, webUrl: string, resultSet: SearchResult[], lastDocId: string = '0'): Promise<SearchResult[]> {
    const requestUrl: string = await this.getRequestUrl(webUrl, logger, args, lastDocId);
    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const searchResult: SearchResult = await request.get(requestOptions);
    resultSet.push(searchResult);

    const rowLimit = args.options.rowLimit ? args.options.rowLimit : 500;

    if (args.options.allResults && searchResult.PrimaryQueryResult.RelevantResults.RowCount === rowLimit) {
      if (this.verbose) {
        await logger.logToStderr(`Processing search query, retrieved ${resultSet.length * searchResult.PrimaryQueryResult.RelevantResults.RowCount} of ${resultSet[0].PrimaryQueryResult.RelevantResults.TotalRows} items...`);
      }

      // When running in allResults mode, paging is done using the DocId property
      // This is a more stable way than using the StartRow property.
      // Explanation: https://learn.microsoft.com/sharepoint/dev/general-development/pagination-for-large-result-sets
      const lastRow = searchResult.PrimaryQueryResult.RelevantResults.Table.Rows[searchResult.PrimaryQueryResult.RelevantResults.RowCount - 1];
      const newLastDocId = lastRow.Cells.filter(cell => cell.Key === 'DocId')[0].Value;

      return this.executeSearchQuery(logger, args, webUrl, resultSet, newLastDocId);
    }

    return resultSet;
  }

  private async getRequestUrl(webUrl: string, logger: Logger, args: CommandArgs, lastDocId: string): Promise<string> {
    // get the list of selected properties
    const selectPropertiesArray: string[] = this.getSelectPropertiesArray(args);

    // get the sort list
    const sortList = this.getSortList(args);

    // transform arg data to query string parameters
    const propertySelectRequestString: string = `&selectproperties='${formatting.encodeQueryParameter(selectPropertiesArray.join(','))}'`;
    const startRowRequestString: string = `&startrow=${args.options.startRow ? args.options.startRow : 0}`;
    const rowLimitRequestString: string = args.options.rowLimit ? `&rowlimit=${args.options.rowLimit}` : (args.options.allResults ? `&rowlimit=500` : '');
    const sourceIdRequestString: string = args.options.sourceId ? `&sourceid='${args.options.sourceId}'` : ``;
    const trimDuplicatesRequestString: string = `&trimduplicates=${args.options.trimDuplicates ? args.options.trimDuplicates : 'false'}`;
    const enableStemmingRequestString: string = `&enablestemming=${typeof (args.options.enableStemming) === 'undefined' ? 'true' : args.options.enableStemming}`;
    const cultureRequestString: string = args.options.culture ? `&culture=${args.options.culture}` : ``;
    const refinementFiltersRequestString: string = args.options.refinementFilters ? `&refinementfilters='${args.options.refinementFilters}'` : ``;
    const queryTemplateRequestString: string = args.options.queryTemplate ? `&querytemplate='${args.options.queryTemplate}'` : ``;
    const sortListRequestString: string = sortList ? `&sortList='${sortList}'` : ``;
    const rankingModelIdRequestString: string = args.options.rankingModelId ? `&rankingmodelid='${args.options.rankingModelId}'` : ``;
    const propertiesRequestString: string = this.getPropertiesRequestString(args);
    const refinersRequestString: string = args.options.refiners ? `&refiners='${args.options.refiners}'` : ``;
    const hiddenConstraintsRequestString: string = args.options.hiddenConstraints ? `&hiddenconstraints='${args.options.hiddenConstraints}'` : ``;
    const clientTypeRequestString: string = args.options.clientType ? `&clienttype='${args.options.clientType}'` : ``;
    const enablePhoneticRequestString: string = typeof (args.options.enablePhonetic) === 'undefined' ? `` : `&enablephonetic=${args.options.enablePhonetic}`;
    const processBestBetsRequestString: string = typeof (args.options.processBestBets) === 'undefined' ? `` : `&processbestbets=${args.options.processBestBets}`;
    const enableQueryRulesRequestString: string = typeof (args.options.enableQueryRules) === 'undefined' ? `` : `&enablequeryrules=${args.options.enableQueryRules}`;
    const processPersonalFavoritesRequestString: string = typeof (args.options.processPersonalFavorites) === 'undefined' ? `` : `&processpersonalfavorites=${args.options.processPersonalFavorites}`;
    const indexDocIdQueryText = args.options.allResults ? ` IndexDocId>${lastDocId}` : '';

    // construct single requestUrl
    const requestUrl = `${webUrl}/_api/search/query?querytext='${args.options.queryText}${indexDocIdQueryText}'`.concat(
      propertySelectRequestString,
      startRowRequestString,
      rowLimitRequestString,
      sourceIdRequestString,
      trimDuplicatesRequestString,
      enableStemmingRequestString,
      cultureRequestString,
      refinementFiltersRequestString,
      queryTemplateRequestString,
      sortListRequestString,
      rankingModelIdRequestString,
      propertiesRequestString,
      refinersRequestString,
      hiddenConstraintsRequestString,
      clientTypeRequestString,
      enablePhoneticRequestString,
      processBestBetsRequestString,
      enableQueryRulesRequestString,
      processPersonalFavoritesRequestString
    );

    if (this.debug) {
      await logger.logToStderr(`RequestURL: ${requestUrl}`);
    }

    return requestUrl;
  }

  private getPropertiesRequestString(args: CommandArgs): string {
    let properties = args.options.properties ? args.options.properties : '';

    if (args.options.sourceName) {
      if (properties && !properties.endsWith(',')) {
        properties += `,`;
      }

      properties += `SourceName:${args.options.sourceName},SourceLevel:SPSite`;
    }

    return properties ? `&properties='${properties}'` : ``;
  }

  private getSelectPropertiesArray(args: CommandArgs): string[] {
    const selectProperties = args.options.selectProperties
      ? args.options.selectProperties.split(',')
      : ['Title', 'OriginalPath'];

    if (args.options.allResults) {
      selectProperties.filter(p => p.toLowerCase() !== 'docid').push('DocId');
    }

    return selectProperties;
  }

  private getSortList(args: CommandArgs): string {
    const sortList = [];
    if (args.options.allResults) {
      sortList.push(formatting.encodeQueryParameter('[DocId]:ascending'));
    }

    if (args.options.sortList) {
      const sortListArray = args.options.sortList.split(',');

      sortListArray.forEach(sortItem => {
        sortList.push(formatting.encodeQueryParameter(sortItem));
      });
    }

    return sortList.join(',');
  }

  private async printResults(logger: Logger, args: CommandArgs, results: SearchResult[]): Promise<void> {
    if (args.options.rawOutput) {
      await logger.log(results);
    }
    else {
      await logger.log(this.getParsedOutput(args, results));
    }

    if (!args.options.output || cli.shouldTrimOutput(args.options.output)) {
      await logger.log('# Rows: ' + results[results.length - 1].PrimaryQueryResult.RelevantResults.TotalRows);
      await logger.log('# Rows (Including duplicates): ' + results[results.length - 1].PrimaryQueryResult.RelevantResults.TotalRowsIncludingDuplicates);
      await logger.log('Elapsed Time: ' + this.getElapsedTime(results));
    }
  }

  private getElapsedTime(results: SearchResult[]): number {
    let totalTime: number = 0;

    results.forEach(result => {
      totalTime += result.ElapsedTime;
    });

    return totalTime;
  }

  private getRowsFromSearchResults(results: SearchResult[]): ResultTableRow[] {
    const searchResultRows: ResultTableRow[] = [];

    for (let i = 0; i < results.length; i++) {
      searchResultRows.push(...results[i].PrimaryQueryResult.RelevantResults.Table.Rows);
    }

    return searchResultRows;
  }

  private getParsedOutput(args: CommandArgs, results: SearchResult[]): any[] {
    const searchResultRows: ResultTableRow[] = this.getRowsFromSearchResults(results);
    const selectProperties: string[] = this.getSelectPropertiesArray(args);
    const outputData: any[] = searchResultRows.map(row => {
      const rowOutput: any = {};

      row.Cells.map(cell => {
        if (selectProperties.filter(prop => { return cell.Key.toUpperCase() === prop.toUpperCase(); }).length > 0) {
          rowOutput[cell.Key] = cell.Value;
        }
      });

      return rowOutput;
    });

    return outputData;
  }
}

export default new SpoSearchCommand();
