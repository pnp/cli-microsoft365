import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { CreateFolderCopyJobsNameConflictBehavior, spo } from '../../../../utils/spo.js';
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
  skipWait?: boolean;
}

class SpoFolderCopyCommand extends SpoCommand {
  private readonly nameConflictBehaviorOptions = ['fail', 'rename'];

  public get name(): string {
    return commands.FOLDER_COPY;
  }

  public get description(): string {
    return 'Copies a folder to another location';
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
        skipWait: !!args.options.skipWait
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

  #initTypes(): void {
    this.types.string.push('webUrl', 'sourceUrl', 'sourceId', 'targetUrl', 'newName', 'nameConflictBehavior');
    this.types.boolean.push('skipWait');
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
        await logger.logToStderr(`Copying folder '${sourcePath}' to '${destinationPath}'...`);
      }

      const copyJobResponse = await spo.createFolderCopyJob(
        args.options.webUrl,
        sourcePath,
        destinationPath,
        {
          nameConflictBehavior: this.getNameConflictBehaviorValue(args.options.nameConflictBehavior),
          newName: args.options.newName,
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
        await logger.logToStderr('Getting information about the destination folder...');
      }

      // Get destination folder data
      const siteRelativeDestinationFolder = '/' + copyJobResult.TargetObjectSiteRelativeUrl.substring(0, copyJobResult.TargetObjectSiteRelativeUrl.lastIndexOf('/'));
      const absoluteWebUrl = destinationPath.substring(0, destinationPath.toLowerCase().lastIndexOf(siteRelativeDestinationFolder.toLowerCase()));

      const requestOptions: CliRequestOptions = {
        url: `${absoluteWebUrl}/_api/Web/GetFolderById('${copyJobResult.TargetObjectUniqueId}')`,
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

  private getNameConflictBehaviorValue(nameConflictBehavior?: string): CreateFolderCopyJobsNameConflictBehavior {
    switch (nameConflictBehavior?.toLowerCase()) {
      case 'fail':
        return CreateFolderCopyJobsNameConflictBehavior.Fail;
      case 'rename':
        return CreateFolderCopyJobsNameConflictBehavior.Rename;
      default:
        return CreateFolderCopyJobsNameConflictBehavior.Fail;
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
      url: `${options.webUrl}/_api/Web/GetFolderById('${options.sourceId}')/ServerRelativePath`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const path = await request.get<{ DecodedUrl: string }>(requestOptions);
    return path.DecodedUrl;
  }

  private getAbsoluteUrl(webUrl: string, url: string): string {
    const result = url.startsWith('https://') ? url : urlUtil.getAbsoluteUrl(webUrl, url);
    return urlUtil.removeTrailingSlashes(result);
  }
}

export default new SpoFolderCopyCommand();
