import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { Localization, ProfileCardProperty, profileCardPropertyNames } from './profileCardProperties.js';
import commands from '../../commands.js';
import { optionsUtils } from '../../../../utils/optionsUtils.js';
import { zod } from '../../../../utils/zod.js';

const customAttributePropertyNames = profileCardPropertyNames.filter(p => p.toLowerCase().startsWith('customattribute'));

export const options = z.looseObject({
  ...globalOptionsZod.shape,
  name: z.string().alias('n'),
  displayName: z.string().optional().alias('d')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TenantPeopleProfileCardPropertySetCommand extends GraphCommand {
  public get name(): string {
    return commands.PEOPLE_PROFILECARDPROPERTY_SET;
  }

  public get description(): string {
    return 'Updates a custom attribute to the profile card property';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .superRefine((opts, ctx) => {
        if (!customAttributePropertyNames.some(p => p.toLowerCase() === opts.name.toLowerCase())) {
          ctx.addIssue({
            code: 'custom',
            message: `${opts.name} is not a valid value for name. Allowed values are ${customAttributePropertyNames.join(', ')}`
          });
        }
      })
      .superRefine((opts, ctx) => {
        const knownKeys = new Set([...Object.keys(globalOptionsZod.shape), 'name', 'displayName']);
        const unknownKeys = Object.keys(opts).filter(k => !knownKeys.has(k));
        const wronglyFormattedOptions = unknownKeys.filter(key => !key.toLowerCase().startsWith('displayname-'));
        if (wronglyFormattedOptions.length > 0) {
          ctx.addIssue({
            code: 'custom',
            message: `Wrong option format detected for the following option(s): ${wronglyFormattedOptions.join(', ')}'. When adding localizations for customAttributes, use the format displayName-<languageTag>.`
          });
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Updating profile card property '${args.options.name}'...`);
      }

      // Get the right casing for the profile card property name
      const profileCardProperty = customAttributePropertyNames.find(p => p.toLowerCase() === args.options.name.toLowerCase());

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/admin/people/profileCardProperties/${profileCardProperty}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          annotations: [
            {
              displayName: args.options.displayName,
              localizations: this.getLocalizations(args.options)
            }
          ]
        }
      };

      const result = await request.patch<ProfileCardProperty>(requestOptions);
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
      this.handleRejectedODataJsonPromise(err);
    }
  }

  /**
   * Transform option to localization object.
   * @example Transform "--displayName-en-US 'Cost center'" to { languageTag: 'en-US', displayName: 'Cost center' }
   */
  private getLocalizations(options: Options): Localization[] {
    const unknownOptions = optionsUtils.getUnknownOptions(options, zod.schemaToOptions(this.schema!));

    const result = Object.keys(unknownOptions).map(o => ({
      languageTag: o.substring(o.indexOf('-') + 1),
      displayName: unknownOptions[o]
    } as Localization));

    return result;
  }
}

export default new TenantPeopleProfileCardPropertySetCommand();