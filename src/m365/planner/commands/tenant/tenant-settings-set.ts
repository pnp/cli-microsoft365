import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import PlannerCommand from '../../../base/PlannerCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  isPlannerAllowed?: boolean;
  allowCalendarSharing?: boolean;
  allowTenantMoveWithDataLoss?: boolean;
  allowTenantMoveWithDataMigration?: boolean;
  allowRosterCreation?: boolean;
  allowPlannerMobilePushNotifications?: boolean;
}

class PlannerTenantSettingsSetCommand extends PlannerCommand {
  public get name(): string {
    return commands.TENANT_SETTINGS_SET;
  }

  public get description(): string {
    return 'Sets Microsoft Planner configuration of the tenant';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        isPlannerAllowed: typeof args.options.isPlannerAllowed !== 'undefined',
        allowCalendarSharing: typeof args.options.allowCalendarSharing !== 'undefined',
        allowTenantMoveWithDataLoss: typeof args.options.allowTenantMoveWithDataLoss !== 'undefined',
        allowTenantMoveWithDataMigration: typeof args.options.allowTenantMoveWithDataMigration !== 'undefined',
        allowRosterCreation: typeof args.options.allowRosterCreation !== 'undefined',
        allowPlannerMobilePushNotifications: typeof args.options.allowPlannerMobilePushNotifications !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--isPlannerAllowed [isPlannerAllowed]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowCalendarSharing [allowCalendarSharing]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowTenantMoveWithDataLoss [allowTenantMoveWithDataLoss]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowTenantMoveWithDataMigration [allowTenantMoveWithDataMigration]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowRosterCreation [allowRosterCreation]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowPlannerMobilePushNotifications [allowPlannerMobilePushNotifications]',
        autocomplete: ['true', 'false']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const optionsArray = [
          args.options.isPlannerAllowed, args.options.allowCalendarSharing, args.options.allowTenantMoveWithDataLoss,
          args.options.allowTenantMoveWithDataMigration, args.options.allowRosterCreation, args.options.allowPlannerMobilePushNotifications
        ];

        if (optionsArray.every(o => typeof o === 'undefined')) {
          return 'You must specify at least one option';
        }

        for (const option of optionsArray) {
          if (typeof option !== 'undefined' && !validation.isValidBoolean(option as any)) {
            return `Value '${option}' is not a valid boolean`;
          }
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/taskAPI/tenantAdminSettings/Settings`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        prefer: 'return=representation'
      },
      responseType: 'json',
      data: {
        isPlannerAllowed: args.options.isPlannerAllowed,
        allowCalendarSharing: args.options.allowCalendarSharing,
        allowTenantMoveWithDataLoss: args.options.allowTenantMoveWithDataLoss,
        allowTenantMoveWithDataMigration: args.options.allowTenantMoveWithDataMigration,
        allowRosterCreation: args.options.allowRosterCreation,
        allowPlannerMobilePushNotifications: args.options.allowPlannerMobilePushNotifications
      }
    };

    request
      .patch(requestOptions)
      .then((result): void => {
        logger.log(result);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new PlannerTenantSettingsSetCommand();