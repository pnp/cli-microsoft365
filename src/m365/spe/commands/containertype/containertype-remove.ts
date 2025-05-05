import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
import { spe } from '../../../../utils/spe.js';
import { cli } from '../../../../cli/cli.js';
import { validation } from '../../../../utils/validation.js';
import request, { CliRequestOptions } from '../../../../request.js';
import config from '../../../../config.js';

const options = globalOptionsZod
  .extend({
    id: zod.alias('i', z.string()
      .refine(id => validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      }))
      .optional()
    ),
    name: zod.alias('n', z.string().optional()),
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpeContainerTypeRemoveCommand extends SpoCommand {

  public get name(): string {
    return commands.CONTAINERTYPE_REMOVE;
  }

  public get description(): string {
    return 'Remove a specific container type';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: z.ZodTypeAny): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.id, options.name].filter(o => o !== undefined).length === 1, {
        message: 'Use one of the following options: id, name.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (!args.options.force) {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove container type ${args.options.id || args.options.name}?` });

      if (!result) {
        return;
      }
    }

    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.verbose);
      const containerTypeId = await this.getContainerTypeId(args.options, spoAdminUrl);
      const formDigestInfo = await spo.ensureFormDigest(spoAdminUrl, logger, undefined, this.debug);

      if (this.verbose) {
        await logger.logToStderr(`Removing container type ${args.options.id || args.options.name}...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': formDigestInfo.FormDigestValue
        },
        responseType: 'json',
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="7" ObjectPathId="6" /><Method Name="RemoveSPOContainerType" Id="8" ObjectPathId="6"><Parameters><Parameter TypeId="{b66ab1ca-fd51-44f9-8cfc-01f5c2a21f99}"><Property Name="ContainerTypeId" Type="Guid">{${containerTypeId}}</Property></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="6" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
      };

      const result = await request.post<any>(requestOptions);
      if (result[0].ErrorInfo) {
        throw result[0].ErrorInfo.ErrorMessage;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getContainerTypeId(options: Options, spoAdminUrl: string): Promise<string> {
    if (options.id) {
      return options.id;
    }

    return spe.getContainerTypeIdByName(spoAdminUrl, options.name!);
  }
}

export default new SpeContainerTypeRemoveCommand();