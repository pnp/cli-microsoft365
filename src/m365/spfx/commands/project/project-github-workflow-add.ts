import fs from 'fs';
import path from 'path';
import yaml from 'yaml';
import { CommandError } from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import { fsUtil } from '../../../../utils/fsUtil.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { workflow } from './DeployWorkflow.js';
import { BaseProjectCommand } from './base-project-command.js';
import { GitHubWorkflow, GitHubWorkflowStep } from './project-github-workflow-model.js';
import { versions } from '../SpfxCompatibilityMatrix.js';
import { spfx } from '../../../../utils/spfx.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name?: string;
  branchName?: string;
  manuallyTrigger?: boolean;
  loginMethod?: string;
  scope?: string;
  skipFeatureDeployment?: boolean;
  siteUrl?: string;
}

class SpfxProjectGithubWorkflowAddCommand extends BaseProjectCommand {
  private static loginMethod: string[] = ['application', 'user'];
  private static scope: string[] = ['tenant', 'sitecollection'];
  public static ERROR_NO_PROJECT_ROOT_FOLDER: number = 1;

  public get name(): string {
    return commands.PROJECT_GITHUB_WORKFLOW_ADD;
  }

  public get description(): string {
    return 'Adds a GitHub workflow for a SharePoint Framework project.';
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
        name: typeof args.options.name !== 'undefined',
        branchName: typeof args.options.branchName !== 'undefined',
        manuallyTrigger: !!args.options.manuallyTrigger,
        loginMethod: typeof args.options.loginMethod !== 'undefined',
        scope: typeof args.options.scope !== 'undefined',
        skipFeatureDeployment: !!args.options.skipFeatureDeployment
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name [name]'
      },
      {
        option: '-b, --branchName [branchName]'
      },
      {
        option: '-m, --manuallyTrigger'
      },
      {
        option: '-l, --loginMethod [loginMethod]',
        autocomplete: SpfxProjectGithubWorkflowAddCommand.loginMethod
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: SpfxProjectGithubWorkflowAddCommand.scope
      },
      {
        option: '-u, --siteUrl [siteUrl]'
      },
      {
        option: '--skipFeatureDeployment'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.scope && args.options.scope === 'sitecollection') {
          if (!args.options.siteUrl) {
            return `siteUrl option has to be defined when scope set to ${args.options.scope}`;
          }

          const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.siteUrl);
          if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
          }
        }

        if (args.options.loginMethod && SpfxProjectGithubWorkflowAddCommand.loginMethod.indexOf(args.options.loginMethod) < 0) {
          return `${args.options.loginMethod} is not a valid login method. Allowed values are ${SpfxProjectGithubWorkflowAddCommand.loginMethod.join(', ')}`;
        }

        if (args.options.scope && SpfxProjectGithubWorkflowAddCommand.scope.indexOf(args.options.scope) < 0) {
          return `${args.options.scope} is not a valid scope. Allowed values are ${SpfxProjectGithubWorkflowAddCommand.scope.join(', ')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    this.projectRootPath = this.getProjectRoot(process.cwd());
    if (this.projectRootPath === null) {
      throw new CommandError(`Couldn't find project root folder`, SpfxProjectGithubWorkflowAddCommand.ERROR_NO_PROJECT_ROOT_FOLDER);
    }

    const solutionPackageJsonFile: string = path.join(this.projectRootPath, 'package.json');
    const packageJson: string = fs.readFileSync(solutionPackageJsonFile, 'utf-8');
    const solutionName = JSON.parse(packageJson).name;

    if (this.debug) {
      await logger.logToStderr(`Adding GitHub workflow in the current SPFx project`);
    }

    try {
      this.updateWorkflow(solutionName, workflow, args.options);
      this.saveWorkflow(workflow);
    }
    catch (error: any) {
      if (error instanceof CommandError) {
        throw error;
      }

      const message = error instanceof Error ? error.message : String(error);
      throw new CommandError(message);
    }
  }

  private saveWorkflow(workflow: GitHubWorkflow): void {
    const githubPath: string = path.join(this.projectRootPath as string, '.github');
    fsUtil.ensureDirectory(githubPath);

    const workflowPath: string = path.join(githubPath, 'workflows');
    fsUtil.ensureDirectory(workflowPath);

    const workflowFile: string = path.join(workflowPath, 'deploy-spfx-solution.yml');
    fs.writeFileSync(path.resolve(workflowFile), yaml.stringify(workflow), 'utf-8');
  }

  private updateWorkflow(solutionName: string, workflow: GitHubWorkflow, options: GlobalOptions): void {
    workflow.name = options.name ? options.name : workflow.name.replace('{{ name }}', solutionName);

    if (options.branchName) {
      workflow.on.push.branches[0] = options.branchName;
    }

    const version = this.getProjectVersion();

    if (!version) {
      throw new CommandError('Unable to determine the version of the current SharePoint Framework project. Could not find the correct version based on @microsoft/generator-sharepoint property in the .yo-rc.json file.');
    }

    const versionRequirements = versions[version];

    if (!versionRequirements) {
      throw new CommandError(`Could not find Node version for ${version} of SharePoint Framework`);
    }

    const nodeVersion: string = spfx.getHighestNodeVersion(versionRequirements.node.range);

    this.assignNodeVersion(workflow, nodeVersion);

    if (options.manuallyTrigger) {
      // eslint-disable-next-line camelcase
      workflow.on.workflow_dispatch = null;
    }

    if (options.skipFeatureDeployment) {
      this.getDeployAction(workflow).with!.SKIP_FEATURE_DEPLOYMENT = true;
    }

    if (options.loginMethod === 'user') {
      const loginAction = this.getLoginAction(workflow);
      loginAction.with = {
        ADMIN_USERNAME: '${{ secrets.ADMIN_USERNAME }}',
        ADMIN_PASSWORD: '${{ secrets.ADMIN_PASSWORD }}'
      };
    }

    if (options.scope === 'sitecollection') {
      const deployAction = this.getDeployAction(workflow);
      deployAction.with!.SCOPE = 'sitecollection';
      deployAction.with!.SITE_COLLECTION_URL = options.siteUrl;
    }

    if (solutionName) {
      const deployAction = this.getDeployAction(workflow);
      deployAction.with!.APP_FILE_PATH = deployAction.with!.APP_FILE_PATH!.replace('{{ solutionName }}', solutionName);
    }
  }

  private assignNodeVersion(workflow: GitHubWorkflow, nodeVersion: string): void {
    workflow.jobs['build-and-deploy'].env.NodeVersion = nodeVersion;
  }

  private getLoginAction(workflow: GitHubWorkflow): GitHubWorkflowStep {
    const steps = this.getWorkFlowSteps(workflow);
    return steps.find(step => step.uses && step.uses.indexOf('action-cli-login') >= 0)!;
  }

  private getDeployAction(workflow: GitHubWorkflow): GitHubWorkflowStep {
    const steps = this.getWorkFlowSteps(workflow);
    return steps.find(step => step.uses && step.uses.indexOf('action-cli-deploy') >= 0)!;
  }

  private getWorkFlowSteps(workflow: GitHubWorkflow): GitHubWorkflowStep[] {
    return workflow.jobs['build-and-deploy'].steps;
  }
}

export default new SpfxProjectGithubWorkflowAddCommand();
