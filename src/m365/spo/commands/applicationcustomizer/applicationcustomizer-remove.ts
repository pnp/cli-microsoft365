import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
import { formatting } from '../../../../utils/formatting.js';
import { CustomAction } from '../../commands/customaction/customaction.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().alias('u'),
  title: z.string().optional().alias('t'),
  id: z.string().optional().alias('i'),
  clientSideComponentId: z.string().optional().alias('c'),
  scope: z.enum(['Site', 'Web', 'All']).optional().alias('s'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoApplicationCustomizerRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.APPLICATIONCUSTOMIZER_REMOVE;
  }

  public get description(): string {
    return 'Removes an application customizer that is added to a site';
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
      .refine(args => [args.id, args.title, args.clientSideComponentId].filter(value => value !== undefined).length === 1, {
        error: `Specify either 'id', 'title', or 'clientSideComponentId'.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'title', 'clientSideComponentId']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.force) {
        return await this.removeApplicationCustomizer(logger, args.options);
      }

      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the application customizer '${args.options.clientSideComponentId || args.options.title || args.options.id}'?` });

      if (result) {
        await this.removeApplicationCustomizer(logger, args.options);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async removeApplicationCustomizer(logger: Logger, options: Options): Promise<void> {
    const applicationCustomizer = await this.getApplicationCustomizer(options);

    if (this.verbose) {
      await logger.logToStderr(`Removing application customizer '${options.clientSideComponentId || options.title || options.id}' from the site '${options.webUrl}'...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/${applicationCustomizer.Scope.toString() === '2' ? 'Site' : 'Web'}/UserCustomActions('${applicationCustomizer.Id}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    await request.delete(requestOptions);
  }

  private async getApplicationCustomizer(options: Options): Promise<CustomAction> {
    const resolvedScope = options.scope || 'All';
    let appCustomizers: CustomAction[] = [];

    if (options.id) {
      const appCustomizer = await spo.getCustomActionById(options.webUrl, options.id, resolvedScope);

      if (appCustomizer) {
        appCustomizers.push(appCustomizer);
      }
    }
    else if (options.title) {
      appCustomizers = await spo.getCustomActions(options.webUrl, resolvedScope, `(Title eq '${formatting.encodeQueryParameter(options.title as string)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`);
    }
    else {
      appCustomizers = await spo.getCustomActions(options.webUrl, resolvedScope, `(ClientSideComponentId eq guid'${options.clientSideComponentId}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`);
    }

    if (appCustomizers.length === 0) {
      throw `No application customizer with ${options.title && `title '${options.title}'` || options.clientSideComponentId && `ClientSideComponentId '${options.clientSideComponentId}'` || options.id && `id '${options.id}'`} found`;
    }

    if (appCustomizers.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', appCustomizers);
      return await cli.handleMultipleResultsFound<CustomAction>(`Multiple application customizer with ${options.title ? `title '${options.title}'` : `ClientSideComponentId '${options.clientSideComponentId}'`} found.`, resultAsKeyValuePair);
    }

    return appCustomizers[0];
  }
}

export default new SpoApplicationCustomizerRemoveCommand();