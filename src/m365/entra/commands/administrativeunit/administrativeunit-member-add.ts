import GlobalOptions from "../../../../GlobalOptions.js";
import { Logger } from "../../../../cli/Logger.js";
import { entraAdministrativeUnit } from "../../../../utils/entraAdministrativeUnit.js";
import { entraGroup } from "../../../../utils/entraGroup.js";
import { entraUser } from "../../../../utils/entraUser.js";
import { validation } from "../../../../utils/validation.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
import request, { CliRequestOptions } from "../../../../request.js";
import { entraDevice } from "../../../../utils/entraDevice.js";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  administrativeUnitId?: string;
  administrativeUnitName?: string;
  userId?: string;
  userName?: string;
  groupId?: string;
  groupName?: string;
  deviceId?: string;
  deviceName?: string;
}

class EntraAdministrativeUnitMemberAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_MEMBER_ADD;
  }

  public get description(): string {
    return 'Adds a member (user, group, device) to an administrative unit';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        deviceId: typeof args.options.deviceId !== 'undefined',
        deviceName: typeof args.options.deviceName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --administrativeUnitId [administrativeUnitId]'
      },
      {
        option: '-n, --administrativeUnitName [administrativeUnitName]'
      },
      {
        option: "--userId [userId]"
      },
      {
        option: "--userName [userName]"
      },
      {
        option: "--groupId [groupId]"
      },
      {
        option: "--groupName [groupName]"
      },
      {
        option: "--deviceId [deviceId]"
      },
      {
        option: "--deviceName [deviceName]"
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.administrativeUnitId && !validation.isValidGuid(args.options.administrativeUnitId as string)) {
          return `${args.options.administrativeUnitId} is not a valid GUID`;
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId as string)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.groupId && !validation.isValidGuid(args.options.groupId as string)) {
          return `${args.options.groupId} is not a valid GUID`;
        }

        if (args.options.deviceId && !validation.isValidGuid(args.options.deviceId as string)) {
          return `${args.options.deviceId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['administrativeUnitId', 'administrativeUnitName'] });
    this.optionSets.push({ options: ['userId', 'userName', 'groupId', 'groupName', 'deviceId', 'deviceName'] });
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