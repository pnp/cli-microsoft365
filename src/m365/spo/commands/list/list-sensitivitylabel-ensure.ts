import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `${e.input} is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  name: z.string().alias('n').optional(),
  id: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .alias('i')
    .optional(),
  listTitle: z.string().alias('t').optional(),
  listId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .alias('l')
    .optional(),
  listUrl: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListSensitivityLabelEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_SENSITIVITYLABEL_ENSURE;
  }

  public get description(): string {
    return 'Applies a default sensitivity label to the specified document library';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.name, opts.id].filter(o => o !== undefined).length === 1, {
        error: 'Specify one of the following options: name, id.',
        params: {
          customCode: 'optionSet',
          options: ['name', 'id']
        }
      })
      .refine(opts => [opts.listId, opts.listTitle, opts.listUrl].filter(o => o !== undefined).length === 1, {
        error: 'Specify one of the following options: listId, listTitle, listUrl.',
        params: {
          customCode: 'optionSet',
          options: ['listId', 'listTitle', 'listUrl']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const sensitivityLabelId: string = await this.getSensitivityLabelId(args, logger);

      if (this.verbose) {
        await logger.logToStderr(`Applying a sensitivity label ${sensitivityLabelId} to the document library...`);
      }

      let requestUrl: string = `${args.options.webUrl}/_api/web`;
      if (args.options.listId) {
        requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
      }
      else if (args.options.listTitle) {
        requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
      }
      else if (args.options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
        requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
      }

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata',
          'if-match': '*'
        },
        data: { 'DefaultSensitivityLabelForLibrary': sensitivityLabelId },
        responseType: 'json'
      };

      await request.patch<any>(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getSensitivityLabelId(args: CommandArgs, logger: Logger): Promise<string> {
    const { id, name } = args.options;

    if (id) {
      return id;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving sensitivity label id of ${name}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels?$filter=name eq '${formatting.encodeQueryParameter(name!)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ value: { id: string }[] }>(requestOptions);
    const sensitivityLabelItem: { id: string } | undefined = res.value[0];

    if (!sensitivityLabelItem) {
      throw Error(`The specified sensitivity label does not exist`);
    }

    return sensitivityLabelItem.id;
  }
}

export default new SpoListSensitivityLabelEnsureCommand();
