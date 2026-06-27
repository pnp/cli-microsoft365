import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
import { formatting } from '../../../../utils/formatting.js';
import { CustomAction } from '../../commands/customaction/customaction.js';
import { cli } from '../../../../cli/cli.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().alias('u'),
  title: z.string().optional().alias('t'),
  id: z.string().optional().alias('i'),
  clientSideComponentId: z.string().optional().alias('c'),
  newTitle: z.string().optional(),
  description: z.string().optional(),
  clientSideComponentProperties: z.string().refine(val => {
    try {
      JSON.parse(val);
      return true;
    }
    catch {
      return false;
    }
  }, {
    error: e => `An error has occurred while parsing clientSideComponentProperties: ${e.input}`
  }).optional().alias('p'),
  hostProperties: z.string().refine(val => {
    try {
      JSON.parse(val);
      return true;
    }
    catch {
      return false;
    }
  }, {
    error: e => `An error has occurred while parsing hostProperties: ${e.input}`
  }).optional(),
  scope: z.enum(['Site', 'Web', 'All']).optional().alias('s')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoApplicationCustomizerSetCommand extends SpoCommand {
  public get name(): string {
    return commands.APPLICATIONCUSTOMIZER_SET;
  }

  public get description(): string {
    return 'Updates an existing Application Customizer on a site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(args => validation.isValidSharePointUrl(args.webUrl) === true, {
        error: () => 'SharePoint Online site URL must be a string.',
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
      })
      .refine(args => args.newTitle !== undefined || args.description !== undefined || args.clientSideComponentProperties !== undefined || args.hostProperties !== undefined, {
        error: `Please specify an option to be updated`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appCustomizer = await this.getAppCustomizerToUpdate(logger, args.options);
      await this.updateAppCustomizer(logger, args.options, appCustomizer);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async updateAppCustomizer(logger: Logger, options: Options, appCustomizer: CustomAction): Promise<void> {
    const { clientSideComponentProperties, hostProperties, webUrl, newTitle, description }: Options = options;

    if (this.verbose) {
      await logger.logToStderr(`Updating application customizer with ID '${appCustomizer.Id}' on the site '${webUrl}'...`);
    }

    const requestBody: any = {
      HostProperties: hostProperties
    };

    if (newTitle) {
      requestBody.Title = newTitle;
    }

    if (description !== undefined) {
      requestBody.Description = description;
    }

    if (clientSideComponentProperties !== undefined) {
      requestBody.ClientSideComponentProperties = clientSideComponentProperties;
    }

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/${appCustomizer.Scope.toString() === '2' ? 'Site' : 'Web'}/UserCustomActions('${appCustomizer.Id}')`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'X-HTTP-Method': 'MERGE'
      },
      data: requestBody,
      responseType: 'json'
    };

    await request.post<CustomAction>(requestOptions);
  }

  private async getAppCustomizerToUpdate(logger: Logger, options: Options): Promise<CustomAction> {
    const { id, webUrl, title, clientSideComponentId, scope }: Options = options;
    const resolvedScope = scope || 'All';

    if (this.verbose) {
      await logger.logToStderr(`Getting application customizer ${title || clientSideComponentId || id} to update...`);
    }

    let appCustomizers: CustomAction[] = [];

    if (id) {
      const appCustomizer = await spo.getCustomActionById(webUrl, id, resolvedScope);
      if (appCustomizer && appCustomizer.Location === 'ClientSideExtension.ApplicationCustomizer') {
        appCustomizers.push(appCustomizer);
      }
    }
    else if (title) {
      appCustomizers = await spo.getCustomActions(webUrl, resolvedScope, `(Title eq '${formatting.encodeQueryParameter(title as string)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`);
    }
    else {
      appCustomizers = await spo.getCustomActions(webUrl, resolvedScope, `(ClientSideComponentId eq guid'${clientSideComponentId}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`);
    }

    if (appCustomizers.length === 0) {
      throw `No application customizer with ${title && `title '${title}'` || clientSideComponentId && `ClientSideComponentId '${clientSideComponentId}'` || id && `id '${id}'`} found`;
    }

    if (appCustomizers.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', appCustomizers);
      return await cli.handleMultipleResultsFound<CustomAction>(`Multiple application customizer with ${title ? `title '${title}'` : `ClientSideComponentId '${clientSideComponentId}'`} found.`, resultAsKeyValuePair);
    }

    return appCustomizers[0];
  }
}

export default new SpoApplicationCustomizerSetCommand();