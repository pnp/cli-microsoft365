import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { profileCardPropertyNames, ProfileCardProperty } from './profileCardProperties.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string().alias('n')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TenantPeopleProfileCardPropertyGetCommand extends GraphCommand {
  public get name(): string {
    return commands.PEOPLE_PROFILECARDPROPERTY_GET;
  }

  public get description(): string {
    return 'Retrieves information about a specific profile card property';
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
            message: `'${opts.name}' is not a valid value for option name. Allowed values are: ${profileCardPropertyNames.join(', ')}.`
          });
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving information about profile card property '${args.options.name}'...`);
      }

      // Get the right casing for the profile card property name
      const profileCardProperty = profileCardPropertyNames.find(p => p.toLowerCase() === args.options.name.toLowerCase());

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/admin/people/profileCardProperties/${profileCardProperty}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const result = await request.get<ProfileCardProperty>(requestOptions);
      let output: any = result;

      // Transform the output to make it more readable
      if (args.options.output && args.options.output !== 'json' && result.annotations.length > 0) {
        output = result.annotations[0].localizations.reduce((acc, curr) => ({
          ...acc,
          ['displayName ' + curr.languageTag]: curr.displayName
        }), {
          ...result,
          displayName: result.annotations[0].displayName
        });

        delete output.annotations;
      }

      await logger.log(output);
    }
    catch (err: any) {
      if (err.response?.status === 404) {
        this.handleError(`Profile card property '${args.options.name}' does not exist.`);
      }

      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TenantPeopleProfileCardPropertyGetCommand();