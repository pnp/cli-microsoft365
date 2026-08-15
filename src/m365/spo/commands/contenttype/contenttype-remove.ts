import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
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
  id: z.string().optional().alias('i'),
  name: z.string().optional().alias('n'),
  force: z.boolean().optional().alias('f')
});

type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoContentTypeRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_REMOVE;
  }

  public get description(): string {
    return 'Deletes site content type';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(opts => [opts.id, opts.name].filter(x => x !== undefined).length === 1, {
      error: `Specify either 'id' or 'name', but not both.`,
      params: {
        customCode: 'optionSet',
        options: ['id', 'name']
      }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let contentTypeId: string = '';

    const contentTypeIdentifierLabel: string = args.options.id ?
      `with id ${args.options.id}` :
      `with name ${args.options.name}`;

    const removeContentType = async (): Promise<void> => {
      try {
        if (this.debug) {
          await logger.logToStderr(`Retrieving information about the content type ${contentTypeIdentifierLabel}...`);
        }

        let contentTypeIdResult: { value: { StringId: string }[] };
        if (args.options.id) {
          contentTypeIdResult = { "value": [{ "StringId": args.options.id }] };
        }
        else {
          if (this.verbose) {
            await logger.logToStderr(`Looking up the ID of content type ${contentTypeIdentifierLabel}...`);
          }

          const requestOptions: CliRequestOptions = {
            url: `${args.options.webUrl}/_api/web/availableContentTypes?$filter=(Name eq '${formatting.encodeQueryParameter(args.options.name as string)}')`,
            headers: {
              accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
          };

          contentTypeIdResult = await request.get<{ value: { StringId: string }[] }>(requestOptions);
        }

        let res: any;
        if (contentTypeIdResult &&
          contentTypeIdResult.value &&
          contentTypeIdResult.value.length > 0) {
          contentTypeId = contentTypeIdResult.value[0].StringId;

          //execute delete operation
          const requestOptions: CliRequestOptions = {
            url: `${args.options.webUrl}/_api/web/contenttypes('${formatting.encodeQueryParameter(contentTypeId)}')`,
            headers: {
              'X-HTTP-Method': 'DELETE',
              'If-Match': '*',
              'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
          };

          res = await request.post<any>(requestOptions);
        }
        else {
          res = { "odata.null": true };
        }

        if (res && res["odata.null"] === true) {
          throw `Content type not found`;
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeContentType();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the content type ${args.options.id || args.options.name}?` });

      if (result) {
        await removeContentType();
      }
    }
  }
}

export default new SpoContentTypeRemoveCommand();