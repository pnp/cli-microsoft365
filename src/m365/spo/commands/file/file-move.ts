import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { CreateFileCopyJobsNameConflictBehavior, spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  sourceUrl: z.string().optional().alias('s'),
  sourceId: z.uuid().optional().alias('i'),
  targetUrl: z.string().alias('t'),
  newName: z.string().optional(),
  nameConflictBehavior: z.enum(['fail', 'replace', 'rename']).optional(),
  withItemPermissions: z.boolean().optional(),
  bypassSharedLock: z.boolean().optional(),
  skipWait: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileMoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_MOVE;
  }

  public get description(): string {
    return 'Moves a file to another location';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.sourceUrl, opts.sourceId].filter(o => o !== undefined).length === 1, {
        error: `Specify 'sourceUrl' or 'sourceId', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['sourceUrl', 'sourceId']
        }
      });
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
        await logger.logToStderr(`Moving file '${sourceServerRelativePath}' to '${args.options.targetUrl}'...`);
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
          includeItemPermissions: !!args.options.withItemPermissions,
          newName: newName,
          operation: 'move'
        }
      );

      if (args.options.skipWait) {
        return;
      }

      if (this.verbose) {
        await logger.logToStderr('Waiting for the move job to complete...');
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

export default new SpoFileMoveCommand();
