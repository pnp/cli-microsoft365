import commands from '../commands';
import request from '../../../request';
import GlobalOptions from '../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../Command';
import SpoCommand from '../../base/SpoCommand';
import Utils from '../../../Utils';
import { SearchResult } from './search/datatypes/SearchResult';
import { ResultTableRow } from './search/datatypes/ResultTableRow';
import { isNumber } from 'util';
import { CommandInstance } from '../../../cli';

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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.selectproperties = typeof args.options.selectProperties !== 'undefined';
    telemetryProps.allResults = args.options.allResults;
    telemetryProps.rowLimit = args.options.rowLimit;
    telemetryProps.sourceId = typeof args.options.sourceId !== 'undefined';
    telemetryProps.trimDuplicates = args.options.trimDuplicates;
    telemetryProps.enableStemming = args.options.enableStemming;
    telemetryProps.culture = args.options.culture;
    telemetryProps.refinementFilters = typeof args.options.refinementFilters !== 'undefined';
    telemetryProps.queryTemplate = typeof args.options.queryTemplate !== 'undefined';
    telemetryProps.sortList = typeof args.options.sortList !== 'undefined';
    telemetryProps.rankingModelId = typeof args.options.rankingModelId !== 'undefined';
    telemetryProps.startRow = typeof args.options.startRow !== 'undefined';
    telemetryProps.properties = typeof args.options.properties !== 'undefined';
    telemetryProps.sourceName = typeof args.options.sourceName !== 'undefined';
    telemetryProps.refiners = typeof args.options.refiners !== 'undefined';
    telemetryProps.webUrl = typeof args.options.webUrl !== 'undefined';
    telemetryProps.hiddenConstraints = typeof args.options.hiddenConstraints !== 'undefined';
    telemetryProps.clientType = typeof args.options.clientType !== 'undefined';
    telemetryProps.enablePhonetic = args.options.enablePhonetic;
    telemetryProps.processBestBets = args.options.processBestBets;
    telemetryProps.enableQueryRules = args.options.enableQueryRules;
    telemetryProps.processPersonalFavorites = args.options.processPersonalFavorites;
    telemetryProps.rawOutput = args.options.rawOutput;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let webUrl: string = '';

      ((): Promise<string> => {
        if (args.options.webUrl) {
          return Promise.resolve(args.options.webUrl);
        }
        else {
          return this.getSpoUrl(cmd, this.debug);
        }
      })()
      .then((_webUrl: string): Promise<SearchResult[]> => {
        webUrl = _webUrl;

        if (this.verbose) {
          cmd.log(`Executing search query '${args.options.queryText}' on site at ${webUrl}...`);
        }

        const startRow = args.options.startRow ? args.options.startRow : 0;

        return this.executeSearchQuery(cmd, args, webUrl, [], startRow);
      })
      .then((results: SearchResult[]) => {
        this.printResults(cmd, args, results);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private executeSearchQuery(cmd: CommandInstance, args: CommandArgs, webUrl: string, resultSet: SearchResult[], startRow: number): Promise<SearchResult[]> {
    return ((): Promise<SearchResult> => {
      const requestUrl: string = this.getRequestUrl(webUrl, cmd, args, startRow);
      const requestOptions: any = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        json: true
      };

      return request.get(requestOptions);
    })()
      .then((searchResult: SearchResult): SearchResult => {
        resultSet.push(searchResult);

        return searchResult;
      })
      .then((searchResult: SearchResult): Promise<SearchResult[]> => {
        if (args.options.allResults) {
          if (startRow + searchResult.PrimaryQueryResult.RelevantResults.RowCount < searchResult.PrimaryQueryResult.RelevantResults.TotalRows) {
            const nextStartRow = startRow + searchResult.PrimaryQueryResult.RelevantResults.RowCount;
            return this.executeSearchQuery(cmd, args, webUrl, resultSet, nextStartRow);
          }
        }
        return new Promise<SearchResult[]>((resolve) => { resolve(resultSet); });
      })
      .then(() => { return resultSet });
  }

  private getRequestUrl(webUrl: string, cmd: CommandInstance, args: CommandArgs, startRow: number): string {
    // get the list of selected properties
    const selectPropertiesArray: string[] = this.getSelectPropertiesArray(args);

    // transform arg data to query string parameters
    const propertySelectRequestString: string = `&selectproperties='${encodeURIComponent(selectPropertiesArray.join(","))}'`
    const startRowRequestString: string = `&startrow=${startRow ? startRow : 0}`;
    const rowLimitRequestString: string = args.options.rowLimit ? `&rowlimit=${args.options.rowLimit}` : ``;
    const sourceIdRequestString: string = args.options.sourceId ? `&sourceid='${args.options.sourceId}'` : ``;
    const trimDuplicatesRequestString: string = `&trimduplicates=${args.options.trimDuplicates ? args.options.trimDuplicates : "false"}`;
    const enableStemmingRequestString: string = `&enablestemming=${typeof (args.options.enableStemming) === 'undefined' ? "true" : args.options.enableStemming}`;
    const cultureRequestString: string = args.options.culture ? `&culture=${args.options.culture}` : ``;
    const refinementFiltersRequestString: string = args.options.refinementFilters ? `&refinementfilters='${args.options.refinementFilters}'` : ``;
    const queryTemplateRequestString: string = args.options.queryTemplate ? `&querytemplate='${args.options.queryTemplate}'` : ``;
    const sortListRequestString: string = args.options.sortList ? `&sortList='${encodeURIComponent(args.options.sortList)}'` : ``;
    const rankingModelIdRequestString: string = args.options.rankingModelId ? `&rankingmodelid='${args.options.rankingModelId}'` : ``;
    const propertiesRequestString: string = this.getPropertiesRequestString(args);
    const refinersRequestString: string = args.options.refiners ? `&refiners='${args.options.refiners}'` : ``;
    const hiddenConstraintsRequestString: string = args.options.hiddenConstraints ? `&hiddenconstraints='${args.options.hiddenConstraints}'` : ``;
    const clientTypeRequestString: string = args.options.clientType ? `&clienttype='${args.options.clientType}'` : ``;
    const enablePhoneticRequestString: string = typeof (args.options.enablePhonetic) === 'undefined' ? `` : `&enablephonetic=${args.options.enablePhonetic}`;
    const processBestBetsRequestString: string = typeof (args.options.processBestBets) === 'undefined' ? `` : `&processbestbets=${args.options.processBestBets}`;
    const enableQueryRulesRequestString: string = typeof (args.options.enableQueryRules) === 'undefined' ? `` : `&enablequeryrules=${args.options.enableQueryRules}`;
    const processPersonalFavoritesRequestString: string = typeof (args.options.processPersonalFavorites) === 'undefined' ? `` : `&processpersonalfavorites=${args.options.processPersonalFavorites}`;

    // construct single requestUrl
    const requestUrl = `${webUrl}/_api/search/query?querytext='${args.options.queryText}'`.concat(
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
      cmd.log(`RequestURL: ${requestUrl}`);
    }

    return requestUrl;
  }

  private getPropertiesRequestString(args: CommandArgs): string {
    let properties = args.options.properties ? args.options.properties : '';

    if (args.options.sourceName) {
      if (properties && !properties.endsWith(",")) {
        properties += `,`;
      }

      properties += `SourceName:${args.options.sourceName},SourceLevel:SPSite`;
    }

    return properties ? `&properties='${properties}'` : ``;
  }

  private getSelectPropertiesArray(args: CommandArgs) {
    return args.options.selectProperties
      ? args.options.selectProperties.split(",")
      : ["Title", "OriginalPath"];
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-q, --queryText <queryText>',
        description: 'Query to be executed in KQL format'
      },
      {
        option: '-p, --selectProperties [selectProperties]',
        description: 'Comma-separated list of properties to retrieve. Will retrieve all properties if not specified and json output is requested.'
      },
      {
        option: '-u, --webUrl [webUrl]',
        description: 'The web against which we want to execute the query. If the parameter is not defined, the query is executed against the web that\'s used when logging in to the SPO environment.'
      },
      {
        option: '--allResults',
        description: 'Set, to get all results of the search query, instead of the number specified by the rowlimit (default: 10)'
      },
      {
        option: '--rowLimit [rowLimit]',
        description: 'The number of rows to be returned. When the \'allResults\' option is used, the specified value will define the size of retrieved batches'
      },
      {
        option: '--sourceId [sourceId]',
        description: 'The identifier (GUID) of the result source to be used to run the query.'
      },
      {
        option: '--trimDuplicates',
        description: 'Set, to remove near duplicate items from the search results.'
      },
      {
        option: '--enableStemming',
        description: 'Set, to enable stemming.'
      },
      {
        option: '--culture [culture]',
        description: 'The locale for the query.'
      },
      {
        option: '--refinementFilters [refinementFilters]',
        description: 'The set of refinement filters used when issuing a refinement query.'
      },
      {
        option: '--queryTemplate [queryTemplate]',
        description: 'A string that contains the text that replaces the query text, as part of a query transformation.'
      },
      {
        option: '--sortList [sortList]',
        description: 'The list of properties by which the search results are ordered.'
      },
      {
        option: '--rankingModelId [rankingModelId]',
        description: 'The ID of the ranking model to use for the query.'
      },
      {
        option: '--startRow [startRow]',
        description: 'The first row that is included in the search results that are returned. You use this parameter when you want to implement paging for search results.'
      },
      {
        option: '--properties [properties]',
        description: 'Additional properties for the query.'
      },
      {
        option: '--sourceName [sourceName]',
        description: 'Specified the name of the result source to be used to run the query.'
      },
      {
        option: '--refiners [refiners]',
        description: 'The set of refiners to return in a search result.'
      },
      {
        option: '--hiddenConstraints [hiddenConstraints]',
        description: 'The additional query terms to append to the query.'
      },
      {
        option: '--clientType [clientType]',
        description: 'The type of the client that issued the query.'
      },
      {
        option: '--enablePhonetic',
        description: 'Set, to use the phonetic forms of the query terms to find matches. (Default = false).'
      },
      {
        option: '--processBestBets',
        description: 'Set, to return best bet results for the query.'
      },
      {
        option: '--enableQueryRules',
        description: 'Set, to enable query rules for the query. '
      },
      {
        option: '--processPersonalFavorites',
        description: 'Set, to return personal favorites with the search results.'
      },
      {
        option: '--rawOutput',
        description: 'Set, to return the unparsed, raw results of the REST call to the search API.'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.sourceId && !Utils.isValidGuid(args.options.sourceId)) {
        return `${args.options.sourceId} is not a valid GUID`;
      }

      if (args.options.rankingModelId && !Utils.isValidGuid(args.options.rankingModelId)) {
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

      if (args.options.culture && !isNumber(args.options.culture)) {
        return `${args.options.culture} is not a valid number`;
      }

      return true;
    };
  }

  private printResults(cmd: CommandInstance, args: CommandArgs, results: SearchResult[]): void {
    if (args.options.rawOutput) {
      cmd.log(results);
    }
    else {
      cmd.log(this.getParsedOutput(args, results));
    }

    if (!args.options.output || args.options.output == 'text') {
      cmd.log("# Rows: " + results[results.length - 1].PrimaryQueryResult.RelevantResults.TotalRows);
      cmd.log("# Rows (Including duplicates): " + results[results.length - 1].PrimaryQueryResult.RelevantResults.TotalRowsIncludingDuplicates);
      cmd.log("Elapsed Time: " + this.getElapsedTime(results));
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
    let searchResultRows: ResultTableRow[] = [];

    for (let i = 0; i < results.length; i++) {
      searchResultRows.push(...results[i].PrimaryQueryResult.RelevantResults.Table.Rows);
    }

    return searchResultRows;
  }

  private getParsedOutput(args: CommandArgs, results: SearchResult[]): any[] {
    const searchResultRows: ResultTableRow[] = this.getRowsFromSearchResults(results);
    const selectProperties: string[] = this.getSelectPropertiesArray(args);
    const outputData: any[] = searchResultRows.map(row => {
      let rowOutput: any = {};

      row.Cells.map(cell => {
        if (selectProperties.filter(prop => { return cell.Key.toUpperCase() === prop.toUpperCase() }).length > 0) {
          rowOutput[cell.Key] = cell.Value;
        }
      })

      return rowOutput;
    });

    return outputData;
  }
}

module.exports = new SpoSearchCommand();
