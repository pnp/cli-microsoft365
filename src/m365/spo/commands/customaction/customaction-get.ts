import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { CustomAction } from './customaction.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(val => validation.isValidGuid(val), {
    error: e => `${e.input} is not valid. Custom action id (Guid) expected.`
  }).optional().alias('i'),
  title: z.string().optional().alias('t'),
  clientSideComponentId: z.string().refine(val => validation.isValidGuid(val), {
    error: e => `${e.input} is not a valid GUID.`
  }).optional().alias('c'),
  webUrl: z.string().refine(val => validation.isValidSharePointUrl(val) === true, {
    error: e => `${e.input} is not a valid SharePoint Online site URL.`
  }).alias('u'),
  scope: z.enum(['Site', 'Web', 'All']).optional().alias('s')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoCustomActionGetCommand extends SpoCommand {
  public get name(): string {
    return commands.CUSTOMACTION_GET;
  }

  public get description(): string {
    return 'Gets details for the specified custom action';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.title, opts.clientSideComponentId].filter(x => x !== undefined).length === 1, {
        message: 'Specify one of id, title or clientSideComponentId',
        params: {
          customCode: 'optionSet',
          options: ['id', 'title', 'clientSideComponentId']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const customAction = await this.getCustomAction(args.options);

      if (customAction) {
        await logger.log({
          ClientSideComponentId: customAction.ClientSideComponentId,
          ClientSideComponentProperties: customAction.ClientSideComponentProperties,
          CommandUIExtension: customAction.CommandUIExtension,
          Description: customAction.Description,
          Group: customAction.Group,
          Id: customAction.Id,
          ImageUrl: customAction.ImageUrl,
          Location: customAction.Location,
          Name: customAction.Name,
          RegistrationId: customAction.RegistrationId,
          RegistrationType: customAction.RegistrationType,
          Rights: JSON.stringify(customAction.Rights),
          Scope: this.humanizeScope(customAction.Scope),
          ScriptBlock: customAction.ScriptBlock,
          ScriptSrc: customAction.ScriptSrc,
          Sequence: customAction.Sequence,
          Title: customAction.Title,
          Url: customAction.Url,
          VersionOfUserCustomAction: customAction.VersionOfUserCustomAction
        });
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private async getCustomAction(options: Options): Promise<CustomAction | undefined> {
    if (options.id) {
      const customAction: CustomAction | undefined = await spo.getCustomActionById(options.webUrl, options.id, options.scope);

      if (!customAction) {
        throw `No user custom action with id '${options.id}' found`;
      }

      return customAction;
    }
    else if (options.title) {
      const customActions: CustomAction[] = await spo.getCustomActions(options.webUrl, options.scope, `Title eq '${formatting.encodeQueryParameter(options.title as string)}'`);

      if (customActions.length === 1) {
        return customActions[0];
      }

      if (customActions.length === 0) {
        throw `No user custom action with title '${options.title}' found`;
      }

      const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', customActions);
      return await cli.handleMultipleResultsFound<CustomAction>(`Multiple user custom actions with title '${options.title}' found.`, resultAsKeyValuePair);
    }
    else {
      const customActions: CustomAction[] = await spo.getCustomActions(options.webUrl, options.scope, `ClientSideComponentId eq guid'${options.clientSideComponentId}'`);

      if (customActions.length === 0) {
        throw `No user custom action with ClientSideComponentId '${options.clientSideComponentId}' found`;
      }

      if (customActions.length > 1) {
        const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', customActions);
        return await cli.handleMultipleResultsFound<CustomAction>(`Multiple user custom actions with ClientSideComponentId '${options.clientSideComponentId}' found.`, resultAsKeyValuePair);
      }

      return customActions[0];
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

export default new SpoCustomActionGetCommand();