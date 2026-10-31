import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { SiteDesign } from './SiteDesign.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(val => validation.isValidGuid(val), { message: 'The value must be a valid GUID.' }).optional().alias('i'),
  title: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteDesignGetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_GET;
  }

  public get description(): string {
    return 'Gets information about the specified site design';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.title].filter(x => x !== undefined).length === 1, {
        message: 'Specify either id or title, but not both.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'title']
        }
      });
  }

  private async getSiteDesignId(args: CommandArgs, spoUrl: string): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const requestOptions: CliRequestOptions = {
      url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response: { value: SiteDesign[] } = await request.post<{ value: SiteDesign[] }>(requestOptions);

    const matchingSiteDesigns: SiteDesign[] = response.value.filter(x => x.Title === args.options.title);

    if (matchingSiteDesigns.length === 0) {
      throw `The specified site design does not exist`;
    }

    if (matchingSiteDesigns.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', matchingSiteDesigns);
      const result = await cli.handleMultipleResultsFound<{ Id: string }>(`Multiple site designs with title '${args.options.title}' found.`, resultAsKeyValuePair);
      return result.Id;
    }

    return matchingSiteDesigns[0].Id;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const siteDesignId: string = await this.getSiteDesignId(args, spoUrl);
      const requestOptions: any = {
        url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata`,
        headers: {
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        },
        data: { id: siteDesignId },
        responseType: 'json'
      };

      const res = await request.post(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteDesignGetCommand();