import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
import Auth from '../../../../Auth.js';

const consumingTenantOverridablesOptions = ['urlTemplate', 'isDiscoverabilityEnabled', 'isSearchEnabled', 'isItemVersioningEnabled', 'itemMajorVersionLimit', 'maxStoragePerContainerInBytes'];

const options = globalOptionsZod
  .extend({
    name: zod.alias('n', z.string()),
    appId: z.string()
      .refine(id => validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      })).optional(),
    billingType: z.enum(['standard', 'trial', 'directToCustomer']).default('standard'),
    consumingTenantOverridables: z.string()
      .refine(values => values.split(',').every(v => consumingTenantOverridablesOptions.includes(v.trim())), values => ({
        message: `'${values}' is not a valid value. Valid options are: ${consumingTenantOverridablesOptions.join(', ')}.`
      })).optional(),
    isDiscoverabilityEnabled: z.boolean().optional(),
    isItemVersioningEnabled: z.boolean().optional(),
    isSearchEnabled: z.boolean().optional(),
    isSharingRestricted: z.boolean().optional(),
    itemMajorVersionLimit: z.number()
      .refine(n => validation.isValidPositiveInteger(n), n => ({
        message: `'${n}' is not a valid positive integer.`
      })).optional(),
    maxStoragePerContainerInBytes: z.number()
      .refine(n => validation.isValidPositiveInteger(n), n => ({
        message: `'${n}' is not a valid positive integer.`
      })).optional(),
    sharingCapability: z.enum(['disabled', 'externalUserSharingOnly', 'externalUserAndGuestSharing', 'existingExternalUserSharingOnly']).optional(),
    urlTemplate: z.string().optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpeContainerTypeAddCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.CONTAINERTYPE_ADD;
  }

  public get description(): string {
    return 'Creates a new container type';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => !options.itemMajorVersionLimit || options.isItemVersioningEnabled !== false, {
        message: `Cannot set itemMajorVersionLimit when isItemVersioningEnabled is false.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Creating a new Container Type for your app with name '${args.options.name}'.`);
      }

      const appId = args.options.appId ?? Auth.connection.appId!;

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/storage/fileStorage/containerTypes`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          name: args.options.name,
          owningAppId: appId,
          billingClassification: args.options.billingType,
          settings: {
            consumingTenantOverridables: args.options.consumingTenantOverridables?.split(',').map(s => s.trim()).join(','),
            isDiscoverabilityEnabled: args.options.isDiscoverabilityEnabled,
            isItemVersioningEnabled: args.options.isItemVersioningEnabled,
            isSearchEnabled: args.options.isSearchEnabled,
            isSharingRestricted: args.options.isSharingRestricted,
            itemMajorVersionLimit: args.options.itemMajorVersionLimit,
            maxStoragePerContainerInBytes: args.options.maxStoragePerContainerInBytes,
            sharingCapability: args.options.sharingCapability,
            urlTemplate: args.options.urlTemplate
          }
        }
      };

      const response = await request.post<any>(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpeContainerTypeAddCommand();