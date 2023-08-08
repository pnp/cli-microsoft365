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
  retainEditorAndModified?: boolean;
  bypassSharedLock?: boolean;
}

class SpoFolderMoveCommand extends SpoCommand {
  private readonly nameConflictBehaviorOptions = ['fail', 'rename'];

  public get name(): string {
    return commands.FOLDER_MOVE;
  }

  public get description(): string {
    return 'Moves a folder to another location';
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
        sourceUrl: typeof args.options.sourceUrl !== 'undefined',
        sourceId: typeof args.options.sourceId !== 'undefined',
        newName: typeof args.options.newName !== 'undefined',
        nameConflictBehavior: typeof args.options.nameConflictBehavior !== 'undefined',
        retainEditorAndModified: !!args.options.retainEditorAndModified,
        bypassSharedLock: !!args.options.bypassSharedLock
      });
    });
  }

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
        option: '--retainEditorAndModified'
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
          return `'${args.options.sourceId}' is not a valid GUID for sourceId.`;
        }

        if (args.options.nameConflictBehavior && this.nameConflictBehaviorOptions.indexOf(args.options.nameConflictBehavior) === -1) {
          return `'${args.options.nameConflictBehavior}' is not a valid value for nameConflictBehavior. Allowed values are: ${this.nameConflictBehaviorOptions.join(', ')}.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['sourceUrl', 'sourceId'] });
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['targetUrl', 'sourceUrl'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const sourcePath = await this.getSourcePath(logger, args.options);

      if (this.verbose) {
        await logger.logToStderr(`Moving folder '${sourcePath}' to '${args.options.targetUrl}'...`);
      }

      const absoluteSourcePath = this.getAbsoluteUrl(args.options.webUrl, sourcePath);
      let absoluteTargetPath = this.getAbsoluteUrl(args.options.webUrl, args.options.targetUrl) + '/';

      if (args.options.newName) {
        absoluteTargetPath += args.options.newName;
      }
      else {
        // Keep the original file name
        absoluteTargetPath += sourcePath.substring(sourcePath.lastIndexOf('/') + 1);
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/SP.MoveCopyUtil.MoveFolderByPath`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          srcPath: {
            DecodedUrl: absoluteSourcePath
          },
          destPath: {
            DecodedUrl: absoluteTargetPath
          },
          options: {
            KeepBoth: args.options.nameConflictBehavior === 'rename',
            ShouldBypassSharedLocks: !!args.options.bypassSharedLock,
            RetainEditorAndModifiedOnMove: !!args.options.retainEditorAndModified
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
      await logger.logToStderr(`Retrieving server-relative path for folder with ID '${options.sourceId}'...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/Web/GetFolderById('${options.sourceId}')?$select=ServerRelativePath`,
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

export default new SpoFolderMoveCommand();
