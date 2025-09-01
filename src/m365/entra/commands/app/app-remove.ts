import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  objectId?: string;
  name?: string;
  force?: boolean;
}

class EntraAppRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_REMOVE;
  }

  public get description(): string {
    return 'Removes an Entra app registration';
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
        appId: typeof args.options.appId !== 'undefined',
        objectId: typeof args.options.objectId !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--appId [appId]' },
      { option: '--objectId [objectId]' },
      { option: '--name [name]' },
      { option: '-f, --force' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId as string)) {
          return `${args.options.appId} is not a valid GUID`;
        }

        if (args.options.objectId && !validation.isValidGuid(args.options.objectId as string)) {
          return `${args.options.objectId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['appId', 'objectId', 'name'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const deleteApp = async (): Promise<void> => {
      try {
        const objectId = await this.getObjectId(args, logger);

        if (this.verbose) {
          await logger.logToStderr(`Deleting Microsoft Entra app ${objectId}...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await deleteApp();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the app?` });

      if (result) {
        await deleteApp();
      }
    }
  }

  private async getObjectId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.objectId) {
      return args.options.objectId;
    }

    const { appId, name } = args.options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appId ? appId : name}...`);
    }

    if (appId) {
      const app = await entraApp.getAppRegistrationByAppId(appId, ['id']);
      return app.id!;
    }
    else {
      const app = await entraApp.getAppRegistrationByAppName(name as string, ['id']);
      return app.id!;
    }
  }
}

export default new EntraAppRemoveCommand();