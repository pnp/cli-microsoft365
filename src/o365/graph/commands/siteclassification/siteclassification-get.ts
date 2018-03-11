import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import GlobalOptions from '../../../../GlobalOptions';

import Utils from '../../../../Utils';
import GraphCommand from '../../GraphCommand';
import { DirectorySettingTemplatesRsp } from './DirectorySettingTemplatesRsp';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
}

class GraphO365SiteClassificationGetCommand extends GraphCommand {
  public get name(): string {
    return `${commands.SITECLASSIFICATION_GET}`;
  }

  public get description(): string {
    return 'Get site classification configuration';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): request.RequestPromise => {
        const requestOptions: any = {
          url: `${auth.service.resource}/beta/settings`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata.metadata=none'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: DirectorySettingTemplatesRsp): void => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }

        if(res.value.length == 0) { 
          // TODO: Handle for the Group.Unified key...
          cmd.log('SiteClassification is not enabled.')
        }
        else{
          cmd.log('SiteClassification is enabled.')
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }


  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to the Microsoft Graph
    using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To get information about a Office 365 Group, you have to first connect to
    the Microsoft Graph using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT}`)}.

  Examples:
  
    Get information about the Office 365 Group with id ${chalk.grey(`1caf7dcd-7e83-4c3a-94f7-932a1299c844`)}
      ${chalk.grey(config.delimiter)} ${this.name} --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844
    `);
  }
}

module.exports = new GraphO365SiteClassificationGetCommand();