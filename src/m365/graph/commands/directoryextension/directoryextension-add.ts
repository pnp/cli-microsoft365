import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { zod } from '../../../../utils/zod.js';
import commands from '../../commands.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { ExtensionProperty } from '@microsoft/microsoft-graph-types';
import { validation } from '../../../../utils/validation.js';
import { entraApp } from '../../../../utils/entraApp.js';

const options = globalOptionsZod
  .extend({
    name: zod.alias('n', z.string()),
    appId: z.string().optional(),
    appObjectId: z.string().optional(),
    appName: z.string().optional(),
    dataType: z.enum(['Binary', 'Boolean', 'DateTime', 'Integer', 'LargeInteger', 'String']),
    targetObjects: z.string().transform((value) => value.split(',').map(String))
      .pipe(z.enum(['User', 'Group', 'Application', 'AdministrativeUnit', 'Device', 'Organization']).array()),
    isMultiValued: z.boolean().optional()
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class GraphDirectoryExtensionAddCommand extends GraphCommand {
  public get name(): string {
    return commands.DIRECTORYEXTENSION_ADD;
  }

  public get description(): string {
    return 'Creates a new directory extension';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => Object.values([options.appId, options.appObjectId, options.appName]).filter(v => typeof v !== 'undefined').length === 1, {
        message: 'Specify either appId, appObjectId or appName, but not multiple'
      })
      .refine(options => (!options.appId && !options.appObjectId && !options.appName) || options.appObjectId || options.appName ||
        (options.appId && validation.isValidGuid(options.appId)), options => ({
        message: `The '${options.appId}' must be a valid GUID`,
        path: ['appId']
      }))
      .refine(options => (!options.appId && !options.appObjectId && !options.appName) || options.appId || options.appName ||
        (options.appObjectId && validation.isValidGuid(options.appObjectId)), options => ({
        message: `The '${options.appObjectId}' must be a valid GUID`,
        path: ['appObjectId']
      }));
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appObjectId = await this.getAppObjectId(args.options);

      if (args.options.verbose) {
        await logger.logToStderr(`Adding directoroy extension to the app with id '${appObjectId}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/applications/${appObjectId}/extensionProperties`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        data: {
          name: args.options.name,
          dataType: args.options.dataType,
          targetObjects: args.options.targetObjects,
          isMultiValued: args.options.isMultiValued
        },
        responseType: 'json'
      };

      const res = await request.post<ExtensionProperty>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppObjectId(options: Options): Promise<string> {
    if (options.appObjectId) {
      return options.appObjectId;
    }

    if (options.appId) {
      return await entraApp.getAppObjectIdFromAppId(options.appId);
    }

    return await entraApp.getAppObjectIdFromAppName(options.appName!);
  }
}

export default new GraphDirectoryExtensionAddCommand();