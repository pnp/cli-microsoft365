import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
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
  id: z.number().optional(),
  name: z.string().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoGroupRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_REMOVE;
  }

  public get description(): string {
    return 'Removes group from specific web';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(o => [o.id, o.name].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'id' or 'name', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeGroup = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing group in web at ${args.options.webUrl}...`);
      }

      try {
        let groupId: number | undefined;
        if (args.options.name) {
          const requestOptions: CliRequestOptions = {
            url: `${args.options.webUrl}/_api/web/sitegroups/GetByName('${args.options.name}')?$select=Id`,
            headers: {
              accept: 'application/json'
            },
            responseType: 'json'
          };
          const group = await request.get<{ Id: number }>(requestOptions);
          groupId = group.Id;
        }
        else {
          groupId = args.options.id;
        }

        const requestUrl = `${args.options.webUrl}/_api/web/sitegroups/RemoveById(${groupId})`;
        const requestOptions: CliRequestOptions = {
          url: requestUrl,
          method: 'POST',
          headers: {
            'content-length': 0,
            'accept': 'application/json'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
        // REST post call doesn't return anything
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeGroup();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the group ${args.options.id || args.options.name} from web ${args.options.webUrl}?` });

      if (result) {
        await removeGroup();
      }
    }
  }
}

export default new SpoGroupRemoveCommand();