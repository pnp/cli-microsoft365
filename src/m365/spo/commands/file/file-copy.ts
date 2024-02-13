import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  sourceUrl?: string;
  sourceId?: string;
  targetUrl: string;
  newName?: string;
  nameConflictBehavior?: string;
  resetAuthorAndCreated?: boolean;
  bypassSharedLock?: boolean;
}

class SpoFileCopyCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_COPY;
  }

  public get description(): string {
    return 'Copies a file to another location';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        sourceUrl: typeof args.options.sourceUrl !== 'undefined',
        sourceId: typeof args.options.sourceId !== 'undefined',
        newName: typeof args.options.newName !== 'undefined',
        nameConflictBehavior: args.options.nameConflictBehavior || false,
        resetAuthorAndCreated: !!args.options.resetAuthorAndCreated,
        bypassSharedLock: !!args.options.bypassSharedLock
      });
    });
  }

  private readonly nameConflictBehaviorOptions = ['fail', 'replace', 'rename'];
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-s, --sourceUrl [sourceUrl]'
      },
      {
        option: '-i, --sourceId [sourceId]'
      },
      {
        option: '-t, --targetUrl <targetUrl>'
      },
      {
        option: '--newName [newName]'
      },
      {
        option: '--nameConflictBehavior [nameConflictBehavior]',
        autocomplete: this.nameConflictBehaviorOptions
      },
      {
        option: '--resetAuthorAndCreated'
      },
      {
        option: '--bypassSharedLock'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.sourceId && !validation.isValidGuid(args.options.sourceId)) {
          return `${args.options.sourceId} is not a valid GUID for sourceId.`;
        }

        if (args.options.nameConflictBehavior && this.nameConflictBehaviorOptions.indexOf(args.options.nameConflictBehavior) === -1) {
          return `${args.options.nameConflictBehavior} is not a valid nameConflictBehavior value. Allowed values: ${this.nameConflictBehaviorOptions.join(', ')}.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['sourceUrl', 'sourceId'] });
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'sourceUrl', 'sourceId', 'targetUrl', 'newName', 'nameConflictBehavior');
    this.types.boolean.push('resetAuthorAndCreated', 'bypassSharedLock');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const sourceServerRelativePath = await this.getSourcePath(logger, args.options);
      const sourcePath = this.getAbsoluteUrl(args.options.webUrl, sourceServerRelativePath);
      let destinationPath = this.getAbsoluteUrl(args.options.webUrl, args.options.targetUrl) + '/';

      if (args.options.newName) {
        destinationPath += args.options.newName;
      }
      else {
        // Keep the original file name
        destinationPath += sourcePath.substring(sourcePath.lastIndexOf('/') + 1);
      }

      if (this.verbose) {
        await logger.logToStderr(`Copying file '${sourcePath}' to '${destinationPath}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/SP.MoveCopyUtil.CopyFileByPath`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          srcPath: {
            DecodedUrl: sourcePath
          },
          destPath: {
            DecodedUrl: destinationPath
          },
          overwrite: args.options.nameConflictBehavior === 'replace',
          options: {
            KeepBoth: args.options.nameConflictBehavior === 'rename',
            ResetAuthorAndCreatedOnCopy: !!args.options.resetAuthorAndCreated,
            ShouldBypassSharedLocks: !!args.options.bypassSharedLock
          }
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getSourcePath(logger: Logger, options: Options): Promise<string> {
    if (options.sourceUrl) {
      return urlUtil.getServerRelativePath(options.webUrl, options.sourceUrl);
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving server-relative path for file with ID '${options.sourceId}'...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/Web/GetFileById('${options.sourceId}')?$select=ServerRelativePath`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const file = await request.get<{ ServerRelativePath: { DecodedUrl: string } }>(requestOptions);
    return file.ServerRelativePath.DecodedUrl;
  }

  private getAbsoluteUrl(webUrl: string, url: string): string {
    return url.startsWith('https://') ? url : urlUtil.getAbsoluteUrl(webUrl, url);
  }
}

export default new SpoFileCopyCommand();
