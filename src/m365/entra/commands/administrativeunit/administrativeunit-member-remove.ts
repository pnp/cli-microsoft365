import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { entraDevice } from '../../../../utils/entraDevice.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.uuid().optional().alias('i'),
  userId: z.uuid().optional(),
  userName: z.string().optional(),
  groupId: z.uuid().optional(),
  groupName: z.string().optional(),
  deviceId: z.uuid().optional(),
  deviceName: z.string().optional(),
  administrativeUnitId: z.uuid().optional(),
  administrativeUnitName: z.string().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAdministrativeUnitMemberRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_MEMBER_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific member (user, group, or device) from an administrative unit';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.administrativeUnitId, options.administrativeUnitName].filter(Boolean).length === 1, {
        error: 'Specify either administrativeUnitId or administrativeUnitName',
        params: {
          customCode: 'optionSet',
          options: ['administrativeUnitId', 'administrativeUnitName']
        }
      })
      .refine(options => [options.id, options.userId, options.userName, options.groupId, options.groupName, options.deviceId, options.deviceName].filter(Boolean).length === 1, {
        error: 'Specify either id, userId, userName, groupId, groupName, deviceId, or deviceName',
        params: {
          customCode: 'optionSet',
          options: ['id', 'userId', 'userName', 'groupId', 'groupName', 'deviceId', 'deviceName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeAdministrativeUnitMember = async (): Promise<void> => {
      let administrativeUnitId = args.options.administrativeUnitId;
      let memberId = args.options.id;

      try {
        if (args.options.administrativeUnitName) {
          if (this.verbose) {
            await logger.logToStderr(`Retrieving Administrative Unit Id...`);
          }

          administrativeUnitId = (await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.administrativeUnitName)).id!;
        }

        if (args.options.userId || args.options.userName) {
          memberId = args.options.userId;

          if (args.options.userName) {
            if (this.verbose) {
              await logger.logToStderr(`Retrieving User Id...`);
            }

            memberId = await entraUser.getUserIdByUpn(args.options.userName);
          }
        }
        else if (args.options.groupId || args.options.groupName) {
          memberId = args.options.groupId;

          if (args.options.groupName) {
            if (this.verbose) {
              await logger.logToStderr(`Retrieving Group Id...`);
            }

            memberId = await entraGroup.getGroupIdByDisplayName(args.options.groupName);
          }
        }
        else if (args.options.deviceId || args.options.deviceName) {
          memberId = args.options.deviceId;

          if (args.options.deviceName) {
            if (this.verbose) {
              await logger.logToStderr(`Retrieving Device Id`);
            }

            memberId = (await entraDevice.getDeviceByDisplayName(args.options.deviceName)).id;
          }
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${memberId}/$ref`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          }
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeAdministrativeUnitMember();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove member '${args.options.id || args.options.userName || args.options.groupName || args.options.deviceName}' from administrative unit '${args.options.administrativeUnitId || args.options.administrativeUnitName}'?` });

      if (result) {
        await removeAdministrativeUnitMember();
      }
    }
  }
}

export default new EntraAdministrativeUnitMemberRemoveCommand();