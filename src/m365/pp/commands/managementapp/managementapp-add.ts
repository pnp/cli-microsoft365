import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { entraApp } from '../../../../utils/entraApp.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  appId: z.string().refine(val => validation.isValidGuid(val), {
    message: 'The value must be a valid GUID.'
  }).optional(),
  objectId: z.string().refine(val => validation.isValidGuid(val), {
    message: 'The value must be a valid GUID.'
  }).optional(),
  name: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpManagementAppAddCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.MANAGEMENTAPP_ADD;
  }

  public get description(): string {
    return 'Registers management application for Power Platform';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.appId, opts.objectId, opts.name].filter(x => x !== undefined).length === 1, {
        message: `Specify either 'appId', 'objectId', or 'name', but not multiple.`,
        params: {
          customCode: 'optionSet',
          options: ['appId', 'objectId', 'name']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appId = await this.getAppId(args);

      const requestOptions: any = {
        // This should be refactored once we implement a PowerPlatform base class as api.bap will differ between envs.
        url: `${this.resource}/providers/Microsoft.BusinessAppPlatform/adminApplications/${appId}?api-version=2020-06-01`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.put(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppId(args: CommandArgs): Promise<string> {
    if (args.options.appId) {
      return args.options.appId;
    }

    const { objectId, name } = args.options;

    if (objectId) {
      const app = await entraApp.getAppRegistrationByObjectId(objectId, ['appId']);
      return app.appId!;
    }
    else {
      const app = await entraApp.getAppRegistrationByAppName(name!, ['appId']);
      return app.appId!;
    }
  }
}

export default new PpManagementAppAddCommand();
