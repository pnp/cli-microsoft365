import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FileProperties } from './FileProperties';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  url?: string;
  id?: string;
  asString?: boolean;
  asListItem?: boolean;
  asFile?: boolean;
  path?: string;
  withPermissions?: boolean;
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
        path: (!(!args.options.path)).toString(),
        withPermissions: !!args.options.withPermissions
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
      },
      {
        option: '--withPermissions'
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
    this.optionSets.push({ options: ['id', 'url'] });
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving file from site ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';
    let options: string = '';

    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileById('${formatting.encodeQueryParameter(args.options.id)}')`;
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

      options += `@f='${formatting.encodeQueryParameter(args.options.url)}'`;
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

    try {
      const file = await request.get<any>(requestOptions);

      if (args.options.asFile && args.options.path) {
        // Not possible to use async/await for this promise
        await new Promise<void>((resolve, reject) => {
          const writer = fs.createWriteStream(args.options.path as string);
          file.data.pipe(writer);

          writer.on('error', err => {
            reject(err);
          });
          writer.on('close', () => {
            const filePath = args.options.path as string;
            if (this.verbose) {
              logger.logToStderr(`File saved to path ${filePath}`);
            }
            return resolve();
          });
        });
      }
      else {
        if (args.options.asString) {
          logger.log(file.toString());
        }
        else {
          const fileProperties: FileProperties = JSON.parse(JSON.stringify(file));
          if (args.options.withPermissions) {
            requestOptions.url = `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${file.ServerRelativeUrl}')/ListItemAllFields/RoleAssignments?$expand=Member,RoleDefinitionBindings`;
            const response = await request.get<{ value: any[] }>(requestOptions);
            response.value.forEach((r: any) => {
              r.RoleDefinitionBindings = formatting.setFriendlyPermissions(r.RoleDefinitionBindings);
            });
            fileProperties.RoleAssignments = response.value;
            if (args.options.asListItem) {
              fileProperties.ListItemAllFields.RoleAssignments = response.value;
            }
          }
          logger.log(args.options.asListItem ? fileProperties.ListItemAllFields : fileProperties);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoFileGetCommand();