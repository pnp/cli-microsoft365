import { Cli,Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
}

class SearchExternalConnectionRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.EXTERNALCONNECTION_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified External Connection from Microsoft Search';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.authorizedAppIds = typeof args.options.authorizedAppIds !== undefined;
    return telemetryProps;
  }

  private getExternalConnectionId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return Promise.resolve(args.options.id);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/external/connections?$filter=name eq '${encodeURIComponent(args.options.name as string)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: [{ id: string }] }>(requestOptions)
      .then(response => {
        const extConn: { id: string } | undefined = response.value[0];

        if (!extConn) {
          return Promise.reject(`The specified connection does not exist in Microsoft Search`);
        }

        if (response.value.length >= 2) {
          return Promise.reject(`Multiple External Connections with name ${args.options.name} found: ${response.value.map(x => x.id)}`);
        }

        return Promise.resolve(extConn.id);
      });
  }


  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeExternalConnection: () => void = (): void => {
      this
        .getExternalConnectionId(args)
        .then((externalConnectionId: string) => {
          const requestOptions: any = {
            url: `${this.resource}/v1.0/external/connections/${encodeURIComponent(externalConnectionId)}`,
            headers: {
              accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
          };

          request
            .delete(requestOptions)
            .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));

        }), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb);
    };

    if (args.options.confirm) {
      removeExternalConnection();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the team ${args.options.teamId}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeExternalConnection();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {

    if (args.options.id && args.options.name) {
      return 'Specify either teamId or teamName, but not both.';
    }

    if (!args.options.id && !args.options.name) {
      return 'Specify teamId or teamName, one is required';
    }

    return true;
  }
}

module.exports = new SearchExternalConnectionRemoveCommand();
