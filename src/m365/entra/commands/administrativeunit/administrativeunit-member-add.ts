import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from "../../../../cli/Logger.js";
import { entraAdministrativeUnit } from "../../../../utils/entraAdministrativeUnit.js";
import { entraGroup } from "../../../../utils/entraGroup.js";
import { entraUser } from "../../../../utils/entraUser.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
import request, { CliRequestOptions } from "../../../../request.js";
import { entraDevice } from "../../../../utils/entraDevice.js";

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  administrativeUnitId: z.uuid().optional().alias('i'),
  administrativeUnitName: z.string().optional().alias('n'),
  userId: z.uuid().optional(),
  userName: z.string().optional(),
  groupId: z.uuid().optional(),
  groupName: z.string().optional(),
  deviceId: z.uuid().optional(),
  deviceName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAdministrativeUnitMemberAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_MEMBER_ADD;
  }

  public get description(): string {
    return 'Adds a member (user, group, or device) to an administrative unit';
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
      .refine(options => [options.userId, options.userName, options.groupId, options.groupName, options.deviceId, options.deviceName].filter(Boolean).length === 1, {
        error: 'Specify either userId, userName, groupId, groupName, deviceId, or deviceName',
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName', 'groupId', 'groupName', 'deviceId', 'deviceName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let administrativeUnitId = args.options.administrativeUnitId;
    let memberType;
    let memberId;

    try {
      if (args.options.administrativeUnitName) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving Administrative Unit Id...`);
        }

        administrativeUnitId = (await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.administrativeUnitName)).id!;
      }

      if (args.options.userId || args.options.userName) {
        memberType = 'users';
        memberId = args.options.userId;

        if (args.options.userName) {
          if (this.verbose) {
            await logger.logToStderr(`Retrieving User Id...`);
          }

          memberId = await entraUser.getUserIdByUpn(args.options.userName);
        }
      }
      else if (args.options.groupId || args.options.groupName) {
        memberType = 'groups';
        memberId = args.options.groupId;

        if (args.options.groupName) {
          if (this.verbose) {
            await logger.logToStderr(`Retrieving Group Id...`);
          }

          memberId = await entraGroup.getGroupIdByDisplayName(args.options.groupName);
        }
      }
      else if (args.options.deviceId || args.options.deviceName) {
        memberType = 'devices';
        memberId = args.options.deviceId;

        if (args.options.deviceName) {
          if (this.verbose) {
            await logger.logToStderr(`Device with name ${args.options.deviceName} retrieved, returned id: ${memberId}`);
          }

          memberId = (await entraDevice.getDeviceByDisplayName(args.options.deviceName)).id;
        }
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/$ref`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        data: {
          "@odata.id": `https://graph.microsoft.com/v1.0/${memberType}/${memberId}`
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraAdministrativeUnitMemberAddCommand();