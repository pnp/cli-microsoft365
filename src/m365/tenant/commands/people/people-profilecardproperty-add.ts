import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { optionsUtils } from '../../../../utils/optionsUtils.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { profileCardPropertyNames } from './profileCardProperties.js';

export const options = z.looseObject({
  ...globalOptionsZod.shape,
  name: z.string().alias('n'),
  displayName: z.string().optional().alias('d')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TenantPeopleProfileCardPropertyAddCommand extends GraphCommand {
  public get name(): string {
    return commands.PEOPLE_PROFILECARDPROPERTY_ADD;
  }

  public get description(): string {
    return 'Adds an additional attribute to the profile card properties';
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
      })
      .refine(
        opts => {
          const propertyName = (opts.name).toLowerCase();
          if (propertyName.startsWith('customattribute') && opts.displayName === undefined) {
            return false;
          }
          return true;
        },
        {
          message: `The option 'displayName' is required when adding customAttributes as profile card properties`,
          params: {
            customCode: 'required'
          }
        }
      )
      .refine(
        opts => {
          const propertyName = (opts.name).toLowerCase();
          if (!propertyName.startsWith('customattribute') && opts.displayName !== undefined) {
            return false;
          }
          return true;
        },
        {
          message: `The option 'displayName' can only be used when adding customAttributes as profile card properties`,
          params: {
            customCode: 'required'
          }
        }
      )
      .superRefine((opts, ctx) => {
        const propertyName = (opts.name).toLowerCase();
        if (!propertyName.startsWith('customattribute')) {
          // For non-custom attributes, check that no unknown options are passed
          const knownKeys = new Set([...Object.keys(globalOptionsZod.shape), 'name', 'displayName']);
          const unknownKeys = Object.keys(opts).filter(k => !knownKeys.has(k));
          if (unknownKeys.length > 0) {
            ctx.addIssue({
              code: 'custom',
              message: `Unknown options like ${unknownKeys.join(', ')} are only supported with customAttributes`
            });
          }
        }
      })
      .superRefine((opts, ctx) => {
        const propertyName = (opts.name).toLowerCase();
        if (propertyName.startsWith('customattribute')) {
          const knownKeys = new Set([...Object.keys(globalOptionsZod.shape), 'name', 'displayName']);
          const unknownKeys = Object.keys(opts).filter(k => !knownKeys.has(k));
          const wronglyFormattedOptions = unknownKeys.filter(key => !key.toLowerCase().startsWith('displayname-'));
          if (wronglyFormattedOptions.length > 0) {
            ctx.addIssue({
              code: 'custom',
              message: `Wrong option format detected for the following option(s): ${wronglyFormattedOptions.join(', ')}'. When adding localizations for customAttributes, use the format displayName-<languageTag>.`
            });
          }
        }
      });
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const directoryPropertyName = profileCardPropertyNames.find(n => n.toLowerCase() === args.options.name.toLowerCase());

    if (this.verbose) {
      await logger.logToStderr(`Adding '${directoryPropertyName}' as a profile card property...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/admin/people/profileCardProperties`,
      headers: {
        'content-type': 'application/json',
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        directoryPropertyName,
        annotations: this.getAnnotations(args.options)
      }
    };

    try {
      const response: any = await request.post(requestOptions);

      // Transform the output to make it more readable
      if (args.options.output && args.options.output !== 'json' && response.annotations.length > 0) {
        const annotation = response.annotations[0];

        response.displayName = annotation.displayName;
        annotation.localizations.forEach((l: { languageTag: string, displayName: string }) => {
          response[`displayName ${l.languageTag}`] = l.displayName;
        });

        delete response.annotations;
      }

      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getAnnotations(options: Options): { displayName: string, localizations?: { languageTag: string, displayName: string }[] }[] {
    if (!options.displayName) {
      return [];
    }

    return [
      {
        displayName: options.displayName!,
        localizations: this.getLocalizations(options)
      }
    ];
  }

  private getLocalizations(options: Options): { languageTag: string, displayName: string }[] {
    const unknownOptions = Object.keys(optionsUtils.getUnknownOptions(options, zod.schemaToOptions(this.schema!)));

    if (unknownOptions.length === 0) {
      return [];
    }

    const localizations: { languageTag: string, displayName: string }[] = [];

    unknownOptions.forEach(key => {
      localizations.push({
        languageTag: key.replace('displayName-', ''),
        displayName: options[key] as string
      });
    });

    return localizations;
  }
}

export default new TenantPeopleProfileCardPropertyAddCommand();