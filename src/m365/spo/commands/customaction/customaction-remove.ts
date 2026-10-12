import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
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
    error: e => `${e.input} is not valid. Custom action Id (GUID) expected.`
  }).optional().alias('i'),
  title: z.string().optional().alias('t'),
  webUrl: z.string().refine(val => validation.isValidSharePointUrl(val) === true, {
    error: e => `${e.input} is not a valid SharePoint Online site URL.`
  }).alias('u'),
  scope: z.enum(['Site', 'Web', 'All']).optional().alias('s'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoCustomActionRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.CUSTOMACTION_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified custom action';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.title].filter(x => x !== undefined).length === 1, {
        message: 'Specify one of id or title',
        params: {
          customCode: 'optionSet',
          options: ['id', 'title']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeCustomAction: () => Promise<void> = async (): Promise<void> => {
      try {
        let customAction: CustomAction | void;
        if (args.options.scope && args.options.scope.toLowerCase() !== "all") {
          customAction = await this.removeScopedCustomAction(args.options);
        }
        else {
          customAction = await this.searchAllScopes(args.options);
        }

        if (this.verbose) {
          if (customAction && customAction["odata.null"] === true) {
            await logger.logToStderr(`Custom action with id ${args.options.id} not found`);
          }
        }
      }
      catch (err: any) {
        this.handleRejectedPromise(err);
      }
    };

    if (args.options.force) {
      await removeCustomAction();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the ${args.options.id} user custom action?` });

      if (result) {
        await removeCustomAction();
      }
    }
  }

  private async getCustomActionId(options: Options): Promise<string> {
    if (options.id) {
      return options.id;
    }

    const customActions = await spo.getCustomActions(options.webUrl, options.scope, `Title eq '${formatting.encodeQueryParameter(options.title as string)}'`);

    if (customActions.length === 1) {
      return customActions[0].Id;
    }

    if (customActions.length === 0) {
      throw `No user custom action with title '${options.title}' found`;
    }

    const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', customActions);
    const result = await cli.handleMultipleResultsFound<CustomAction>(`Multiple user custom actions with title '${options.title}' found.`, resultAsKeyValuePair);
    return result.Id;
  }

  private async removeScopedCustomAction(options: Options): Promise<CustomAction | undefined> {
    const customActionId = await this.getCustomActionId(options);

    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/${options.scope}/UserCustomActions('${formatting.encodeQueryParameter(customActionId)}')')`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'X-HTTP-Method': 'DELETE'
      },
      responseType: 'json'
    };

    return await request.post(requestOptions);
  }

  /**
   * Remove request with `web` scope is send first. 
   * If custom action not found then 
   * another get request is send with `site` scope.
   */
  private async searchAllScopes(options: Options): Promise<CustomAction | undefined> {
    options.scope = "Web";
    const webResult = await this.removeScopedCustomAction(options);
    if (!webResult) {
      return webResult;
    }

    options.scope = "Site";
    const siteResult = await this.removeScopedCustomAction(options);

    return siteResult;
  }
}

export default new SpoCustomActionRemoveCommand();