import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { CustomAction } from '../customaction/customaction.js';
import { cli } from '../../../../cli/cli.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().alias('u'),
  title: z.string().optional().alias('t'),
  id: z.string().optional().alias('i'),
  clientSideComponentId: z.string().optional().alias('c'),
  scope: z.enum(['All', 'Site', 'Web']).optional().alias('s'),
  clientSideComponentProperties: z.boolean().optional().alias('p')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoApplicationCustomizerGetCommand extends SpoCommand {
  public get name(): string {
    return commands.APPLICATIONCUSTOMIZER_GET;
  }

  public get description(): string {
    return 'Gets an application customizer that is added to a site.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(args => validation.isValidSharePointUrl(args.webUrl) === true, {
        error: e => validation.isValidSharePointUrl((e.input as Options).webUrl) as string,
        path: ['webUrl']
      })
      .refine(args => !args.id || validation.isValidGuid(args.id), {
        error: () => 'id is not a valid GUID',
        path: ['id']
      })
      .refine(args => !args.clientSideComponentId || validation.isValidGuid(args.clientSideComponentId), {
        error: () => 'clientSideComponentId is not a valid GUID',
        path: ['clientSideComponentId']
      })
      .refine(args => [args.title, args.id, args.clientSideComponentId].filter(value => value !== undefined).length === 1, {
        error: `Specify either 'title', 'id', or 'clientSideComponentId'.`,
        params: {
          customCode: 'optionSet',
          options: ['title', 'id', 'clientSideComponentId']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const customAction = await this.getCustomAction(args.options);

      if (!args.options.clientSideComponentProperties) {
        await logger.log({
          ...customAction,
          Scope: this.humanizeScope(customAction.Scope)
        });
      }
      else {
        const properties = formatting.tryParseJson(customAction.ClientSideComponentProperties);
        await logger.log(properties);
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private async getCustomAction(options: Options): Promise<CustomAction> {
    if (options.id) {
      const customAction = await spo.getCustomActionById(options.webUrl, options.id, options.scope);

      if (!customAction || (customAction && customAction.Location !== 'ClientSideExtension.ApplicationCustomizer')) {
        throw `No application customizer with id '${options.id}' found`;
      }

      return customAction;
    }

    const filter = options.title ? `Title eq '${formatting.encodeQueryParameter(options.title as string)}'` : `ClientSideComponentId eq guid'${formatting.encodeQueryParameter(options.clientSideComponentId!)}'`;
    const customActions = await spo.getCustomActions(options.webUrl, options.scope, `${filter} and Location eq 'ClientSideExtension.ApplicationCustomizer'`);

    if (customActions.length === 1) {
      return customActions[0];
    }

    const identifier = options.title ? `title '${options.title}'` : `Client Side Component Id '${options.clientSideComponentId}'`;
    if (customActions.length === 0) {
      throw `No application customizer with ${identifier} found`;
    }
    else {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', customActions);
      return await cli.handleMultipleResultsFound<CustomAction>(`Multiple application customizers with ${identifier} found.`, resultAsKeyValuePair);
    }
  }

  private humanizeScope(scope: number): string {
    switch (scope) {
      case 2:
        return "Site";
      case 3:
        return "Web";
    }

    return `${scope}`;
  }
}

export default new SpoApplicationCustomizerGetCommand();