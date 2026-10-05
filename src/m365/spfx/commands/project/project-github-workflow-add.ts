import fs from 'fs';
import path from 'path';
import yaml from 'yaml';
import { z } from 'zod';
import { globalOptionsZod, CommandError } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { fsUtil } from '../../../../utils/fsUtil.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { workflow } from './DeployWorkflow.js';
import { BaseProjectCommand } from './base-project-command.js';
import { GitHubWorkflow, GitHubWorkflowStep } from './project-github-workflow-model.js';
import { Project } from './project-model/index.js';
import { versions } from '../SpfxCompatibilityMatrix.js';
import { spfx } from '../../../../utils/spfx.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string().alias('n').optional(),
  branchName: z.string().alias('b').optional(),
  manuallyTrigger: z.boolean().alias('m').optional(),
  loginMethod: z.enum(['application', 'user']).alias('l').optional(),
  scope: z.enum(['tenant', 'sitecollection']).alias('s').optional(),
  siteUrl: z.string().alias('u').optional(),
  skipFeatureDeployment: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpfxProjectGithubWorkflowAddCommand extends BaseProjectCommand {
  public static ERROR_NO_PROJECT_ROOT_FOLDER: number = 1;

  public get name(): string {
    return commands.PROJECT_GITHUB_WORKFLOW_ADD;
  }

  public get description(): string {
    return 'Adds a GitHub workflow for a SharePoint Framework project.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => !opts.scope || opts.scope !== 'sitecollection' || opts.siteUrl, {
        error: `siteUrl option has to be defined when scope set to sitecollection`,
        params: {
          customCode: 'required'
        }
      })
      .refine(opts => !opts.scope || opts.scope !== 'sitecollection' || validation.isValidSharePointUrl(opts.siteUrl!) === true, {
        error: `The specified siteUrl is not a valid SharePoint Online site URL.`,
        params: {
          customCode: 'required'
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    this.projectRootPath = this.getProjectRoot(process.cwd());
    if (this.projectRootPath === null) {
      throw new CommandError(`Couldn't find project root folder`, SpfxProjectGithubWorkflowAddCommand.ERROR_NO_PROJECT_ROOT_FOLDER);
    }

    try {
      const project: Project = { path: this.projectRootPath };
      this.readAndParseJsonFile(path.join(this.projectRootPath, 'package.json'), project, 'packageJson');
      this.readAndParseJsonFile(path.join(this.projectRootPath, 'config', 'package-solution.json'), project, 'packageSolutionJson');

      const solutionName = project.packageJson!.name!;
      const sppkgPath = (project.packageSolutionJson as any)?.paths?.zippedPackage;

      if (this.debug) {
        await logger.logToStderr(`Adding GitHub workflow in the current SPFx project`);
      }

      const workflowToSave: GitHubWorkflow = structuredClone(workflow);

      this.updateWorkflow(solutionName, sppkgPath, workflowToSave, args.options);
      this.saveWorkflow(workflowToSave);
    }
    catch (error: any) {
      this.handleError(error);
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

  private updateWorkflow(solutionName: string, sppkgPath: string | undefined, workflow: GitHubWorkflow, options: Options): void {
    workflow.name = options.name ? options.name : workflow.name.replace('{{ name }}', solutionName);

    if (options.branchName) {
      workflow.on.push.branches[0] = options.branchName;
    }

    const version = this.getProjectVersion();

    if (!version) {
      throw 'Unable to determine the version of the current SharePoint Framework project. Could not find the correct version based on the version property in the .yo-rc.json file.';
    }

    const versionRequirements = versions[version];

    if (!versionRequirements) {
      throw `Could not find Node version for version '${version}' of SharePoint Framework.`;
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

    if (sppkgPath) {
      const deployAction = this.getDeployAction(workflow);
      deployAction.with!.APP_FILE_PATH = deployAction.with!.APP_FILE_PATH!.replace('{{ sppkgPath }}', sppkgPath);
    }

    if (versionRequirements.heft === undefined) {
      const buildAndPackageStep = this.getBuildAndPackageStep(workflow);
      buildAndPackageStep.run = `gulp bundle --ship\ngulp package-solution --ship\n`;
      buildAndPackageStep.name = 'Bundle & Package';
    }
  }

  private assignNodeVersion(workflow: GitHubWorkflow, nodeVersion: string): void {
    workflow.jobs['build-and-deploy'].env.NodeVersion = nodeVersion;
  }

  private getBuildAndPackageStep(workflow: GitHubWorkflow): GitHubWorkflowStep {
    const steps = this.getWorkFlowSteps(workflow);
    return steps.find(step => step.run && step.name === 'Build & Package')!;
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
