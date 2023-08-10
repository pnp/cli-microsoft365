import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

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
  bypassSharedLock?: boolean;
  resetAuthorAndCreated?: boolean;
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
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        sourceUrl: typeof args.options.sourceUrl !== 'undefined',
        sourceId: typeof args.options.sourceId !== 'undefined',
        newName: typeof args.options.newName !== 'undefined',
        nameConflictBehavior: args.options.nameConflictBehavior || false,
        bypassSharedLock: !!args.options.bypassSharedLock,
        resetAuthorAndCreated: !!args.options.resetAuthorAndCreated
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
        option: '--bypassSharedLock'
      },
      {
        option: '--resetAuthorAndCreated'
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

        if (args.options.sourceId) {
          if (!validation.isValidGuid(args.options.sourceId)) {
            return `${args.options.sourceId} is not a valid GUID`;
          }
        }

        if (args.options.nameConflictBehavior && this.nameConflictBehaviorOptions.indexOf(args.options.nameConflictBehavior) === -1) {
          return `${args.options.nameConflictBehavior} is not a valid nameConflictBehavior value. Allowed values: ${this.nameConflictBehaviorOptions.join(', ')}`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['sourceUrl', 'sourceId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let fileServerRelativePath: string = "";
      if (args.options.sourceId) {
        const requestUrl: string = `${args.options.webUrl}/_api/web/GetFileById('${args.options.sourceId}')?$select=ServerRelativeUrl`;
        const requestOptions: CliRequestOptions = {
          url: requestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const res = await request.get<{ ServerRelativeUrl: string }>(requestOptions);
        fileServerRelativePath = res.ServerRelativeUrl;
      }

      const serverRelativePath: string = args.options.sourceUrl ? urlUtil.getServerRelativePath(args.options.webUrl, args.options.sourceUrl) : fileServerRelativePath;
      const sourcePath = this.getAbsoluteUrl(args.options.webUrl, serverRelativePath);
      let destinationPath = this.getAbsoluteUrl(args.options.webUrl, args.options.targetUrl) + '/';

      if (args.options.newName) {
        destinationPath += args.options.newName;
      }
      else {
        // Keep the original file name
        destinationPath += sourcePath.substring(sourcePath.lastIndexOf('/') + 1);
      }

      if (this.verbose) {
        logger.logToStderr(`Copying file '${sourcePath}' to '${destinationPath}'...`);
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

  private getAbsoluteUrl(webUrl: string, url: string): string {
    return url.startsWith('https://') ? url : urlUtil.getAbsoluteUrl(webUrl, url);
  }
}

module.exports = new SpoFileCopyCommand();
