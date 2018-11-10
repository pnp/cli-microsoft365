import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { Auth } from '../../../../Auth';
import Utils from '../../../../Utils';
import { SearchResult } from './datatypes/SearchResult';
import { ResultTableRow } from './datatypes/ResultTableRow';
import { isNumber } from 'util';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  query: string;
  rowLimit?:number;
  selectProperties?:string;
  webUrl?:string;
  allResults?:boolean;
  sourceId?:string;
  trimDuplicates?:boolean;
  enableStemming?:boolean;
  culture?:number;
  refinementFilters?:string;
  queryTemplate?:string;
  sortList?:string;
  rankingModelId?:string;
  startRow?:number;
  properties?:string;
  sourceName?:string;
  refiners?:string;
  hiddenConstraints?:string;
  clientType?:string;
  enablePhonetic?:boolean;
  processBestBets?:boolean;
  enableQueryRules?:boolean;
  processPersonalFavorites?:boolean;
}

class SearchCommand extends SpoCommand {
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
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const webUrl = args.options.webUrl ? args.options.webUrl : auth.site.url;
    const resource: string = Auth.getResourceFromUrl(webUrl);

    if (this.debug) {
      cmd.log("Calling Search API on = " + auth.site.url);
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): string => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        if (this.verbose) {
          cmd.log(`Executing search query '${args.options.query}' on site at ${webUrl}...`);
        }

        return accessToken;
      })
      .then((accessToken:string):Promise<SearchResult[]> => {
        const startRow = args.options.startRow ? args.options.startRow : 0;

        return this.executeSearchQuery(cmd,args,accessToken,webUrl,[],startRow);
      })
      .then((results:SearchResult[]) => {
        this.printResults(cmd,args,results);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private executeSearchQuery(cmd:CommandInstance,args:CommandArgs,accessToken:string,webUrl:string,resultSet:SearchResult[],startRow:number):Promise<SearchResult[]> {
    return (():request.RequestPromise => { 
        const requestUrl:string = this.getRequestUrl(webUrl,cmd,args,startRow);
        const requestOptions: any = {
          url: requestUrl,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })()
      .then((searchResult: SearchResult): SearchResult => {
        if (this.debug) {
          cmd.log(`${searchResult.PrimaryQueryResult.RelevantResults.TotalRowsIncludingDuplicates} Results found (including duplicates) :`);
          cmd.log('');
          cmd.log(searchResult);
          cmd.log('');
        }

        resultSet.push(searchResult);

        return searchResult;
      })
      .then((searchResult:SearchResult):Promise<SearchResult[]> => {
        if(args.options.allResults) {
          if(startRow + searchResult.PrimaryQueryResult.RelevantResults.RowCount < searchResult.PrimaryQueryResult.RelevantResults.TotalRows) {
            const nextStartRow = startRow + searchResult.PrimaryQueryResult.RelevantResults.RowCount;
            return this.executeSearchQuery(cmd,args,accessToken,webUrl,resultSet,nextStartRow);
          }
        } 
        return new Promise<SearchResult[]>((resolve) => { resolve(resultSet); });
      })
      .then(() => { return resultSet });
  }

  private getRequestUrl(webUrl:string, cmd:CommandInstance, args: CommandArgs,startRow:number): string {
    //Get arg data
    const selectPropertiesArray: string[] = this.getSelectPropertiesArray(args);

    //Transform arg data to requeststrings
    const propertySelectRequestString: string = selectPropertiesArray.length > 0 ?
        `&selectproperties='${encodeURIComponent(selectPropertiesArray.join(","))}'` :
        ``;
    const startRowRequestString = `&startrow=${startRow ? startRow : 0}`;
    const rowLimitRequestString = args.options.rowLimit ? `&rowlimit=${args.options.rowLimit}` : ``;
    const sourceIdRequestString = args.options.sourceId ? `&sourceid='${args.options.sourceId}'` : ``;
    const trimDuplicatesRequestString = `&trimduplicates=${args.options.trimDuplicates ? args.options.trimDuplicates : "false"}`;
    const enableStemmingRequestString = `&enablestemming=${typeof(args.options.enableStemming) === 'undefined' ? "true" : args.options.enableStemming}`;
    const cultureRequestString = args.options.culture ? `&culture=${args.options.culture}` : ``;
    const refinementFiltersRequestString = args.options.refinementFilters ? `&refinementfilters='${args.options.refinementFilters}'` : ``;
    const queryTemplateRequestString = args.options.queryTemplate ? `&querytemplate='${args.options.queryTemplate}'` : ``;
    const sortListRequestString = args.options.sortList ? `&sortList='${encodeURIComponent(args.options.sortList)}'` : ``;
    const rankingModelIdRequestString = args.options.rankingModelId ? `&rankingmodelid='${args.options.rankingModelId}'` : ``;
    const propertiesRequestString = this.getPropertiesRequestString(args);
    const refinersRequestString = args.options.refiners ? `&refiners='${args.options.refiners}'` : ``;
    const hiddenConstraintsRequestString = args.options.hiddenConstraints ? `&hiddenconstraints='${args.options.hiddenConstraints}'` : ``;
    const clientTypeRequestString = args.options.clientType ? `&clienttype='${args.options.clientType}'` : ``;
    const enablePhoneticRequestString = typeof(args.options.enablePhonetic) === 'undefined' ? `` : `&enablephonetic=${args.options.enablePhonetic}`;    
    const processBestBetsRequestString = typeof(args.options.processBestBets) === 'undefined' ? `` : `&processbestbets=${args.options.processBestBets}`;
    const enableQueryRulesRequestString = typeof(args.options.enableQueryRules) === 'undefined' ? `` : `&enablequeryrules=${args.options.enableQueryRules}`;
    const processPersonalFavoritesRequestString = typeof(args.options.processPersonalFavorites) === 'undefined' ? `` : `&processpersonalfavorites=${args.options.processPersonalFavorites}`;

    //Construct single requestUrl
    const requestUrl = `${webUrl}/_api/search/query?querytext='${args.options.query}'`.concat(
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

    if(this.debug) {
      cmd.log(`RequestURL: ${requestUrl}`);
    }
    return requestUrl;
  }

  private getPropertiesRequestString(args: CommandArgs):string {
    let properties = args.options.properties ? args.options.properties : '';
    if(args.options.sourceName) {
      if(properties && !properties.endsWith(",")) { properties += `,`; }
      properties += `SourceName:${args.options.sourceName},SourceLevel:SPSite`;
    }
    return properties ? `&properties='${properties}'` : ``;
  }

  private getSelectPropertiesArray(args: CommandArgs) {
    return args.options.selectProperties 
      ? args.options.selectProperties.split(",")
      : (!args.options.output || args.options.output === "text") 
        ? ["Title","OriginalPath"]
        : [];
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-q, --query <query>',
        description: 'Query to be executed in KQL format'
      },
      {
        option: '-p, --selectProperties [selectProperties]',
        description: 'Comma separated list of properties to retrieve. Will retrieve all properties if not specified and json output is requested.'
      },
      {
        option: '-u, --webUrl [webUrl]',
        description: 'The web against which we want to execute the query. If the parameter is not defined, the query is executed against the web that\'s used when logging in to the SPO environment.'
      },
      {
        option: '--allResults',
        description: 'Set, to get all results of the search query, not only the amount specified by the rowlimit (default: 10)'
      },
      {
        option: '--rowLimit [rowLimit]',
        description: 'Sets the number of rows to be returned. When the \'allResults\' parameter is enabled, it will determines the size of the batches being retrieved'
      },
      {
        option: '--sourceId [sourceId]',
        description: 'Specifies the identifier (GUID) of the result source to be used to run the query.'
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
        option:'--sortList [sortList]',
        description: 'The list of properties by which the search results are ordered.'
      },
      {
        option:'--rankingModelId [rankingModelId]',
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
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.query) {
        return 'Required parameter query missing';
      }
      if (args.options.sourceId && !Utils.isValidGuid(args.options.sourceId)) {
        return `${args.options.sourceId} is not a valid GUID`;
      }
      if (args.options.rankingModelId && !Utils.isValidGuid(args.options.rankingModelId)) {
        return `${args.options.rankingModelId} is not a valid GUID`;
      }
      if(args.options.sortList && !/^([a-z0-9_]+:(ascending|descending))(,([a-z0-9_]+:(ascending|descending)))*$/gi.test(args.options.sortList)) {
        return `sortlist parameter value '${args.options.sortList}' does not match the required pattern (=comma separated list of '<property>:(ascending|descending)'-pattern)`;
      }
      if(args.options.rowLimit && !isNumber(args.options.rowLimit)) {
        return `${args.options.rowLimit} is not a valid number`;
      }      
      if(args.options.startRow && !isNumber(args.options.startRow)) {
        return `${args.options.startRow} is not a valid number`;
      }      
      if(args.options.culture && !isNumber(args.options.culture)) {
        return `${args.options.culture} is not a valid number`;
      }
      return true;
    };
  }

  private printResults(cmd:CommandInstance,args:CommandArgs,results:SearchResult[]) {
    if(results.length === 1) {
      if(args.options.output === 'json') {
        cmd.log(results[0]);
      } else {
        cmd.log(this.getTextOutput(args,results[0].PrimaryQueryResult.RelevantResults.Table.Rows));
      }
    } else {
      if(args.options.output === 'json') {
        cmd.log(results);
      } else {
        let allRows:ResultTableRow[] = [];
        for(let i = 0;i < results.length;i++) {
          allRows.push(...results[i].PrimaryQueryResult.RelevantResults.Table.Rows);
        }
        cmd.log(this.getTextOutput(args,allRows));
      }
    }

    if(!args.options.output || args.options.output == 'text') {
      cmd.log("# Rows: "+results[results.length-1].PrimaryQueryResult.RelevantResults.TotalRows);
      cmd.log("# Rows (Including duplicates): "+results[results.length-1].PrimaryQueryResult.RelevantResults.TotalRowsIncludingDuplicates);
      cmd.log("Elapsed Time: "+this.getElapsedTime(results));
    }
  }

  private getElapsedTime(results:SearchResult[]) {
    var totalTime:number = 0;
    results.forEach(result => {
      totalTime += result.ElapsedTime;
    });
    return totalTime;
  }

  private getTextOutput(args: CommandArgs,searchResultRows: ResultTableRow[]) {
    const selectProperties = this.getSelectPropertiesArray(args);
    const outputData = searchResultRows.map(row => {
      var rowOutput:any = {};

      row.Cells.map(cell => { 
        if(selectProperties.filter(prop => { return cell.Key.toUpperCase() === prop.toUpperCase() }).length > 0) {
          rowOutput[cell.Key] = cell.Value;
        }
      })

      return rowOutput;
    });
    return outputData;
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site,
      using the ${chalk.blue(commands.LOGIN)} command.
  
  Remarks:
  
    To execute a search query you have to first log in to SharePoint using the
    ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Execute search query to retrieve all Document Sets (ContentTypeId = '${chalk.grey('0x0120D520')}') for the english locale
      ${chalk.grey(config.delimiter)} ${commands.SEARCH} --query 'ContentTypeId:0x0120D520' --culture 1033

    Retrieve all documents. For each document, retrieve the Path, Author and FileType.
      ${chalk.grey(config.delimiter)} ${commands.SEARCH} --query 'IsDocument:1' --selectProperties 'Path,Author,FileType' --allResults
    
    Return the top 50 items of which the title starts with 'Marketing' while trimming duplicates.
      ${chalk.grey(config.delimiter)} ${commands.SEARCH} --query 'Title:Marketing*' --rowLimit=50 --trimDuplicates

    Return only items from a specific resultsource (using the source id).
      ${chalk.grey(config.delimiter)} ${commands.SEARCH} --query '*' --sourceId 6e71030e-5e16-4406-9bff-9c1829843083
      `);
  }
}

module.exports = new SearchCommand();