import { isNumber } from 'util';
import { Logger } from '../../../cli';
import {
  CommandOption
} from '../../../Command';
import GlobalOptions from '../../../GlobalOptions';
import request from '../../../request';
import { spo, validation } from '../../../utils';
import SpoCommand from '../../base/SpoCommand';
import commands from '../commands';
import { ResultTableRow } from './search/datatypes/ResultTableRow';
import { SearchResult } from './search/datatypes/SearchResult';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let webUrl: string = '';

    ((): Promise<string> => {
      if (args.options.webUrl) {
        return Promise.resolve(args.options.webUrl);
      }
      else {
        return spo.getSpoUrl(logger, this.debug);
      }
    })()
      .then((_webUrl: string): Promise<SearchResult[]> => {
        webUrl = _webUrl;

        if (this.verbose) {
          logger.logToStderr(`Executing search query '${args.options.queryText}' on site at ${webUrl}...`);
        }

        const startRow = args.options.startRow ? args.options.startRow : 0;

        return this.executeSearchQuery(logger, args, webUrl, [], startRow);
      })
      .then((results: SearchResult[]) => {
        this.printResults(logger, args, results);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private executeSearchQuery(logger: Logger, args: CommandArgs, webUrl: string, resultSet: SearchResult[], startRow: number): Promise<SearchResult[]> {
    return ((): Promise<SearchResult> => {
      const requestUrl: string = this.getRequestUrl(webUrl, logger, args, startRow);
      const requestOptions: any = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
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
            return this.executeSearchQuery(logger, args, webUrl, resultSet, nextStartRow);
          }
        }
        return new Promise<SearchResult[]>((resolve) => { resolve(resultSet); });
      })
      .then(() => resultSet);
  }

  private getRequestUrl(webUrl: string, logger: Logger, args: CommandArgs, startRow: number): string {
    // get the list of selected properties
    const selectPropertiesArray: string[] = this.getSelectPropertiesArray(args);

    // transform arg data to query string parameters
    const propertySelectRequestString: string = `&selectproperties='${encodeURIComponent(selectPropertiesArray.join(","))}'`;
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
      logger.logToStderr(`RequestURL: ${requestUrl}`);
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
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
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

    if (args.options.culture && !isNumber(args.options.culture)) {
      return `${args.options.culture} is not a valid number`;
    }

    return true;
  }

  private printResults(logger: Logger, args: CommandArgs, results: SearchResult[]): void {
    if (args.options.rawOutput) {
      logger.log(results);
    }
    else {
      logger.log(this.getParsedOutput(args, results));
    }

    if (!args.options.output || args.options.output === 'text') {
      logger.log("# Rows: " + results[results.length - 1].PrimaryQueryResult.RelevantResults.TotalRows);
      logger.log("# Rows (Including duplicates): " + results[results.length - 1].PrimaryQueryResult.RelevantResults.TotalRowsIncludingDuplicates);
      logger.log("Elapsed Time: " + this.getElapsedTime(results));
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

module.exports = new SpoSearchCommand();
