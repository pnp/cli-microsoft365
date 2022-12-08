import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import commands from '../../commands';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogScope?: string;
  appCatalogUrl?: string;
  id?: string;
  name?: string;
  skipFeatureDeployment?: boolean;
}

class SpoAppDeployCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_DEPLOY;
  }

  public get description(): string {
    return 'Deploys the specified app in the specified app catalog';
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
        id: (!(!args.options.id)).toString(),
        name: (!(!args.options.name)).toString(),
        appCatalogUrl: (!(!args.options.appCatalogUrl)).toString(),
        skipFeatureDeployment: args.options.skipFeatureDeployment || false,
        appCatalogScope: (!(!args.options.appCatalogScope)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '-u, --appCatalogUrl [appCatalogUrl]'
      },
      {
        option: '--skipFeatureDeployment'
      },
      {
        option: '-s, --appCatalogScope [appCatalogScope]',
        autocomplete: ['tenant', 'sitecollection']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        // verify either 'tenant' or 'sitecollection' specified if scope provided
        if (args.options.appCatalogScope) {
          const testScope: string = args.options.appCatalogScope.toLowerCase();
          if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
            return `appCatalogScope must be either 'tenant' or 'sitecollection'`;
          }

          if (testScope === 'sitecollection' && !args.options.appCatalogUrl) {
            return `You must specify appCatalogUrl when the appCatalogScope is sitecollection`;
          }
        }

        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.appCatalogUrl) {
          return validation.isValidSharePointUrl(args.options.appCatalogUrl);
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'name'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let appId: string = '';
    const scope: string = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';
    let appCatalogUrl: string = '';

    try {
      const spoUrl = await spo.getSpoUrl(logger, this.debug);
      appCatalogUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);

      let res: { UniqueId: string };
      if (args.options.id) {
        if (this.verbose) {
          logger.logToStderr(`Using the specified app id ${args.options.id}`);
        }

        res = { UniqueId: args.options.id };
      }
      else {
        if (this.verbose) {
          logger.logToStderr(`Looking up app id for app named ${args.options.name}...`);
        }

        const requestOptions: any = {
          url: `${appCatalogUrl}/_api/web/getfolderbyserverrelativeurl('AppCatalog')/files('${args.options.name}')?$select=UniqueId`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        res = await request.get(requestOptions);
      }

      appId = res.UniqueId;

      if (this.verbose) {
        logger.logToStderr(`Deploying app...`);
      }

      const requestOptions: any = {
        url: `${appCatalogUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${appId}')/deploy`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata;charset=utf-8'
        },
        data: { 'skipFeatureDeployment': args.options.skipFeatureDeployment || false },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoAppDeployCommand();