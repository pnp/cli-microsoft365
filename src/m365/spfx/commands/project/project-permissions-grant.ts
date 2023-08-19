import fs from 'fs';
import path from 'path';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { CommandError } from '../../../../Command.js';
import spoServicePrincipalGrantAddCommand, { Options as SpoServicePrincipalGrantAddCommandOptions } from '../../../spo/commands/serviceprincipal/serviceprincipal-grant-add.js';
import commands from '../../commands.js';
import { BaseProjectCommand } from './base-project-command.js';
import { WebApiPermissionRequests } from './WebApiPermissionRequests.js';

class SpfxProjectPermissionSGrantCommand extends BaseProjectCommand {
  public static ERROR_NO_PROJECT_ROOT_FOLDER: number = 1;

  public get name(): string {
    return commands.PROJECT_PERMISSIONS_GRANT;
  }

  public get description(): string {
    return 'Grant API permissions defined in the current SPFx project';
  }

  constructor() {
    super();
  }

  public async commandAction(logger: Logger): Promise<void> {
    this.projectRootPath = this.getProjectRoot(process.cwd());
    if (this.projectRootPath === null) {
      throw new CommandError(`Couldn't find project root folder`, SpfxProjectPermissionSGrantCommand.ERROR_NO_PROJECT_ROOT_FOLDER);
    }

    if (this.debug) {
      await logger.logToStderr(`Granting API permissions defined in the current SPFx project`);
    }

    try {
      const webApiPermissionsRequest: Array<WebApiPermissionRequests> = this.getWebApiPermissionRequest(path.join(this.projectRootPath, 'config', 'package-solution.json'));
      for (const permission of webApiPermissionsRequest) {
        const options: SpoServicePrincipalGrantAddCommandOptions = {
          resource: permission.resource,
          scope: permission.scope,
          output: 'json',
          debug: this.debug,
          verbose: this.verbose
        };

        let output = null;
        try {
          output = await Cli.executeCommandWithOutput(spoServicePrincipalGrantAddCommand as Command, { options: { ...options, _: [] } });
        }
        catch (err: any) {
          if (err.error && err.error.message.indexOf('already exists') > -1) {
            await this.warn(logger, err.error.message);
            continue;
          }
          else {
            throw err;
          }
        }
        const getGrantOutput = JSON.parse(output!.stdout);
        await logger.log(getGrantOutput);
      }
    }
    catch (error: any) {
      throw new CommandError(error);
    }
  }

  private getWebApiPermissionRequest(filePath: string): Array<WebApiPermissionRequests> {
    if (!fs.existsSync(filePath)) {
      throw (`The package-solution.json file could not be found`);
    }

    const existingContent: string = fs.readFileSync(filePath, 'utf-8');
    const solutionContent = JSON.parse(existingContent);

    return solutionContent.solution.webApiPermissionRequests;
  }
}

export default new SpfxProjectPermissionSGrantCommand();