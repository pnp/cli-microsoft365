import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import { AxiosRequestConfig } from 'axios';
import { validation } from '../../../../utils';
import request from '../../../../request';
import PlannerCommand from '../../../base/PlannerCommand';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';

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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.isPlannerAllowed = typeof args.options.isPlannerAllowed !== 'undefined';
    telemetryProps.allowCalendarSharing = typeof args.options.allowCalendarSharing !== 'undefined';
    telemetryProps.allowTenantMoveWithDataLoss = typeof args.options.allowTenantMoveWithDataLoss !== 'undefined';
    telemetryProps.allowTenantMoveWithDataMigration = typeof args.options.allowTenantMoveWithDataMigration !== 'undefined';
    telemetryProps.allowRosterCreation = typeof args.options.allowRosterCreation !== 'undefined';
    telemetryProps.allowPlannerMobilePushNotifications = typeof args.options.allowPlannerMobilePushNotifications !== 'undefined';
    return telemetryProps;
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
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
}

module.exports = new PlannerTenantSettingsSetCommand();