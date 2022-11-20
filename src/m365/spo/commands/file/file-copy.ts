import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { AxiosRequestConfig } from 'axios';
import { urlUtil } from '../../../../utils/urlUtil';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  sourceUrl: string;
  targetUrl: string;
  newName?: string;
  nameConflictBehavior?: string;
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
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        newName: typeof args.options.newName !== 'undefined',
        nameConflictBehavior: args.options.nameConflictBehavior || false,
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
        option: '-s, --sourceUrl <sourceUrl>'
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

        if (args.options.nameConflictBehavior && this.nameConflictBehaviorOptions.indexOf(args.options.nameConflictBehavior) === -1) {
          return `${args.options.nameConflictBehavior} is not a valid nameConflictBehavior value. Allowed values: ${this.nameConflictBehaviorOptions.join('|')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Copying file '${args.options.sourceUrl}' to '${args.options.targetUrl}'.`);
      }

      const sourcePath = this.getAbsoluteUrl(args.options, args.options.sourceUrl);
      let destinationPath = this.getAbsoluteUrl(args.options, args.options.targetUrl) + '/';

      if (args.options.newName) {
        destinationPath += args.options.newName;
      }
      else {
        // Keep the original file name
        destinationPath += sourcePath.substring(sourcePath.lastIndexOf('/') + 1);
      }

      const requestOptions: AxiosRequestConfig = {
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

  /** Ensure the URL is an absolute URL. */
  private getAbsoluteUrl(options: Options, url: string): string {
    return url.startsWith('https://') ? url : urlUtil.getAbsoluteUrl(options.webUrl, url);
  }
}

module.exports = new SpoFileCopyCommand();
