import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  scope?: string;
}

class SpoCustomActionListCommand extends SpoCommand {
  public get name(): string {
    return commands.CUSTOMACTION_LIST;
  }

  public get description(): string {
    return 'Lists all user custom actions at the given scope';
  }

  public defaultProperties(): string[] | undefined {
    return ['Name', 'Location', 'Scope', 'Id'];
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
        scope: args.options.scope || 'All'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: ['Site', 'Web', 'All']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (typeof isValidUrl === 'string') {
          return isValidUrl;
        }

        if (args.options.scope) {
          if (args.options.scope !== 'Site' &&
            args.options.scope !== 'Web' &&
            args.options.scope !== 'All') {
            return `${args.options.scope} is not a valid custom action scope. Allowed values are Site|Web|All`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const scope: string = args.options.scope ? args.options.scope : 'All';

      if (this.debug) {
        await logger.logToStderr(`Attempt to get custom actions list with scope: ${scope}`);
        await logger.logToStderr('');
      }

      const customActions = await spo.getCustomActions(args.options.webUrl, args.options.scope);

      if (customActions.length === 0) {
        if (this.verbose) {
          await logger.logToStderr(`Custom actions not found`);
        }
      }
      else {
        if (args.options.output !== 'json') {
          customActions.forEach(a => a.Scope = this.humanizeScope(a.Scope) as any);
        }

        await logger.log(customActions);
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private humanizeScope(scope: number): string {
    switch (scope) {
      case 2:
        return "Site";
      case 3:
        return "Web";
    }

    return `${scope}`;
  }
}

export default new SpoCustomActionListCommand();