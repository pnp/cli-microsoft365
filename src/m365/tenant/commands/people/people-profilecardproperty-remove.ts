import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { profileCardPropertyNames } from './profileCardProperties.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string().alias('n'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TenantPeopleProfileCardPropertyRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.PEOPLE_PROFILECARDPROPERTY_REMOVE;
  }

  public get description(): string {
    return 'Removes an additional attribute from the profile card properties';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .superRefine((opts, ctx) => {
        if (!profileCardPropertyNames.some(p => p.toLowerCase() === opts.name.toLowerCase())) {
          ctx.addIssue({
            code: 'custom',
            message: `${opts.name} is not a valid value for name. Allowed values are ${profileCardPropertyNames.join(', ')}`
          });
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const directoryPropertyName = profileCardPropertyNames.find(n => n.toLowerCase() === args.options.name.toLowerCase());

    const removeProfileCardProperty = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing '${directoryPropertyName}' as a profile card property...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/admin/people/profileCardProperties/${directoryPropertyName}`,
        headers: {
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      try {
        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeProfileCardProperty();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the profile card property '${directoryPropertyName}'?` });

      if (result) {
        await removeProfileCardProperty();
      }
    }
  }
}

export default new TenantPeopleProfileCardPropertyRemoveCommand();