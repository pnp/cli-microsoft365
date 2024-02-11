import GlobalOptions from '../../../../GlobalOptions.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { entraDevice } from '../../../../utils/entraDevice.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  id?: string;
  userId?: string;
  userName?: string;
  groupId?: string;
  groupName?: string;
  deviceId?: string;
  deviceName?: string;
  administrativeUnitId?: string;
  administrativeUnitName?: string;
  force?: boolean;
}

class EntraAdministrativeUnitMemberRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_MEMBER_REMOVE;
  }

  public get description(): string {
    return 'Remove a specific member (user, group, or device) from an administrative unit';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTelemetry();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        deviceId: typeof args.options.deviceId !== 'undefined',
        deviceName: typeof args.options.deviceName !== 'undefined',
        administrativeUnitId: typeof args.options.administrativeUnitId !== 'undefined',
        administrativeUnitName: typeof args.options.administrativeUnitName !== 'undefined',
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
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
      },
      {
        option: '--administrativeUnitId [administrativeUnitId]'
      },
      {
        option: '--administrativeUnitName [administrativeUnitName]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.administrativeUnitId && !validation.isValidGuid(args.options.administrativeUnitId as string)) {
          return `${args.options.administrativeUnitId} is not a valid GUID`;
        }

        if (args.options.id && !validation.isValidGuid(args.options.id as string)) {
          return `${args.options.id} is not a valid GUID`;
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
    this.optionSets.push({ options: ['id', 'userId', 'userName', 'groupId', 'groupName', 'deviceId', 'deviceName'] });
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