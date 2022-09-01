import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FileProperties } from './FileProperties';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  url?: string;
  id?: string;
  asString?: boolean;
  asListItem?: boolean;
  asFile?: boolean;
  path?: string;
}

class SpoFileGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_GET;
  }

  public get description(): string {
    return 'Gets information about the specified file';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: (!(!args.options.id)).toString(),
        url: (!(!args.options.url)).toString(),
        asString: args.options.asString || false,
        asListItem: args.options.asListItem || false,
        asFile: args.options.asFile || false,
        path: (!(!args.options.path)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-w, --webUrl <webUrl>'
      },
      {
        option: '-u, --url [url]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '--asString'
      },
      {
        option: '--asListItem'
      },
      {
        option: '--asFile'
      },
      {
        option: '-p, --path [path]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }
    
        if (args.options.id) {
          if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
          }
        }
    
        if (args.options.asFile && !args.options.path) {
          return 'The path should be specified when the --asFile option is used';
        }
    
        if (args.options.path && !fs.existsSync(path.dirname(args.options.path))) {
          return 'Specified path where to save the file does not exits';
        }
    
        if (args.options.asFile) {
          if (args.options.asListItem || args.options.asString) {
            return 'Specify to retrieve the file either as file, list item or string but not multiple';
          }
        }
    
        if (args.options.asListItem) {
          if (args.options.asFile || args.options.asString) {
            return 'Specify to retrieve the file either as file, list item or string but not multiple';
          }
        }
    
        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'url']);
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving file from site ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';
    let options: string = '';

    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileById('${encodeURIComponent(args.options.id)}')`;
    }
    else if (args.options.url) {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl=@f)`;
    }

    if (args.options.asListItem) {
      options = '?$expand=ListItemAllFields';
    }
    else if (args.options.asFile || args.options.asString) {
      options = '/$value';
    }

    if (args.options.url) {
      if (options.indexOf('?') < 0) {
        options += '?';
      }
      else {
        options += '&';
      }

      options += `@f='${encodeURIComponent(args.options.url)}'`;
    }

    const requestOptions: any = {
      url: requestUrl + options,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      // Set responseType to arraybuffer, otherwise binary data will be encoded
      // to utf8 and binary data is corrupt
      responseType: args.options.asFile ? 'stream' : 'json'
    };

    if (args.options.asFile && args.options.path) {
      request
        .get<any>(requestOptions)
        .then((file: any): Promise<string> => {
          return new Promise((resolve, reject) => {
            const writer = fs.createWriteStream(args.options.path as string);

            file.data.pipe(writer);

            writer.on('error', err => {
              reject(err);
            });
            writer.on('close', () => {
              resolve(args.options.path as string);
            });
          });
        })
        .then((file: string): void => {
          if (this.verbose) {
            logger.logToStderr(`File saved to path ${file}`);
          }
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }
    else {
      request
        .get<string>(requestOptions)
        .then((file: string): void => {
          if (args.options.asString) {
            logger.log(file.toString());
          }
          else {
            const fileProperties: FileProperties = JSON.parse(JSON.stringify(file));
            logger.log(args.options.asListItem ? fileProperties.ListItemAllFields : fileProperties);
          }

          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }

  }
}

module.exports = new SpoFileGetCommand();
