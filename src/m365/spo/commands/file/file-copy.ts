import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { CreateFileCopyJobsNameConflictBehavior, spo } from '../../../../utils/spo.js';
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
  bypassSharedLock?: boolean;
  ignoreVersionHistory?: boolean;
  skipWait?: boolean;
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
        nameConflictBehavior: typeof args.options.nameConflictBehavior !== 'undefined',
        ignoreVersionHistory: !!args.options.ignoreVersionHistory,
        bypassSharedLock: !!args.options.bypassSharedLock,
        skipWait: !!args.options.skipWait
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
        option: '--ignoreVersionHistory'
      },
      {
        option: '--skipWait'
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

        if (args.options.nameConflictBehavior && !this.nameConflictBehaviorOptions.includes(args.options.nameConflictBehavior.toLowerCase())) {
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
    this.types.boolean.push('bypassSharedLock', 'ignoreVersionHistory', 'skipWait');
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['targetUrl', 'sourceUrl'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const sourceServerRelativePath = await this.getSourcePath(logger, args.options);
      const sourcePath = this.getAbsoluteUrl(args.options.webUrl, sourceServerRelativePath);
      const destinationPath = this.getAbsoluteUrl(args.options.webUrl, args.options.targetUrl);

      if (this.verbose) {
        await logger.logToStderr(`Copying file '${sourceServerRelativePath}' to '${args.options.targetUrl}'...`);
      }

      let newName = args.options.newName;
      // Add original file extension if not provided
      if (newName && !newName.includes('.')) {
        newName += sourceServerRelativePath.substring(sourceServerRelativePath.lastIndexOf('.'));
      }

      const copyJobResponse = await spo.createFileCopyJob(
        args.options.webUrl,
        sourcePath,
        destinationPath,
        {
          nameConflictBehavior: this.getNameConflictBehaviorValue(args.options.nameConflictBehavior),
          bypassSharedLock: !!args.options.bypassSharedLock,
          ignoreVersionHistory: !!args.options.ignoreVersionHistory,
          newName: newName,
          operation: 'copy'
        }
      );

      if (args.options.skipWait) {
        return;
      }

      if (this.verbose) {
        await logger.logToStderr('Waiting for the copy job to complete...');
      }

      const copyJobResult = await spo.getCopyJobResult(args.options.webUrl, copyJobResponse);

      if (this.verbose) {
        await logger.logToStderr('Getting information about the destination file...');
      }

      // Get destination file data
      const siteRelativeDestinationFolder = '/' + copyJobResult.TargetObjectSiteRelativeUrl.substring(0, copyJobResult.TargetObjectSiteRelativeUrl.lastIndexOf('/'));
      const absoluteWebUrl = destinationPath.substring(0, destinationPath.toLowerCase().lastIndexOf(siteRelativeDestinationFolder.toLowerCase()));

      const requestOptions: CliRequestOptions = {
        url: `${absoluteWebUrl}/_api/Web/GetFileById('${copyJobResult.TargetObjectUniqueId}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const destinationFile = await request.get<any>(requestOptions);
      await logger.log(destinationFile);
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
      url: `${options.webUrl}/_api/Web/GetFileById('${options.sourceId}')/ServerRelativePath`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const file = await request.get<{ DecodedUrl: string }>(requestOptions);
    return file.DecodedUrl;
  }

  private getNameConflictBehaviorValue(nameConflictBehavior?: string): CreateFileCopyJobsNameConflictBehavior {
    switch (nameConflictBehavior?.toLowerCase()) {
      case 'fail':
        return CreateFileCopyJobsNameConflictBehavior.Fail;
      case 'replace':
        return CreateFileCopyJobsNameConflictBehavior.Replace;
      case 'rename':
        return CreateFileCopyJobsNameConflictBehavior.Rename;
      default:
        return CreateFileCopyJobsNameConflictBehavior.Fail;
    }
  }

  private getAbsoluteUrl(webUrl: string, url: string): string {
    const result = url.startsWith('https://') ? url : urlUtil.getAbsoluteUrl(webUrl, url);
    return urlUtil.removeTrailingSlashes(result);
  }
}

export default new SpoFileCopyCommand();
