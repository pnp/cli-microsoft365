import { AzureDevOpsPipeline } from "./project-azuredevops-pipeline-model";
import { GitHubWorkflow } from "./project-github-workflow-model";

export const workflow: GitHubWorkflow = {
  name: "Deploy Solution {{ name }}",
  on: {
    push: {
      branches: [
        "main"
      ]
    }
  },
  jobs: {
    "build-and-deploy": {
      "runs-on": "ubuntu-latest",
      env: {
        NodeVersion: ""
      },
      steps: [
        {
          name: "Checkout",
          uses: "actions/checkout@v4"
        },
        {
          name: "Use Node.js",
          uses: "actions/setup-node@v4",
          with: {
            "node-version": "${{ env.NodeVersion }}"
          }
        },
        {
          name: "Run npm ci",
          run: "npm ci"
        },
        {
          name: "Bundle & Package",
          run: "gulp bundle --ship\ngulp package-solution --ship\n"
        },
        {
          name: "CLI for Microsoft 365 Login",
          uses: "pnp/action-cli-login@v2.2.4",
          with: {
            "CERTIFICATE_ENCODED": "${{ secrets.CERTIFICATE_ENCODED }}",
            "CERTIFICATE_PASSWORD": "${{ secrets.CERTIFICATE_PASSWORD }}",
            "APP_ID": "${{ secrets.APP_ID }}",
            "TENANT": "${{ secrets.TENANT_ID }}"
          }
        },
        {
          name: "CLI for Microsoft 365 Deploy App",
          uses: "pnp/action-cli-deploy@v4.0.0",
          with: {
            "APP_FILE_PATH": "sharepoint/solution/{{ solutionName }}.sppkg",
            "SKIP_FEATURE_DEPLOYMENT": false,
            "OVERWRITE": true
          }
        }
      ]
    }
  }
};

export const pipeline: AzureDevOpsPipeline = {
  name: "Deploy Solution",
  trigger: {
    branches: {
      include: [
        "main"
      ]
    }
  },
  pool: {
    vmImage: "ubuntu-latest"
  },
  variables: [
    {
      name: "CertificateBase64Encoded",
      value: ""
    },
    {
      name: "CertificateSecureFileId",
      value: ""
    },
    {
      name: "CertificatePassword",
      value: ""
    },
    {
      name: "EntraAppId",
      value: ""
    },
    {
      name: "UserName",
      value: ""
    },
    {
      name: "Password",
      value: ""
    },
    {
      name: "TenantId",
      value: ""
    },
    {
      name: "SharePointBaseUrl",
      value: ""
    },
    {
      name: "PackageName",
      value: ""
    },
    {
      name: "SiteAppCatalogUrl",
      value: ""
    },
    {
      name: "NodeVersion",
      value: ""
    }
  ],
  stages: [
    {
      stage: "Build_and_Deploy",
      jobs: [
        {
          job: "Build_and_Deploy",
          steps: [
            {
              task: "NodeTool@0",
              displayName: "Use Node.js",
              inputs: {
                versionSpec: "$(NodeVersion)"
              }
            },
            {
              task: "Npm@1",
              displayName: "Run npm install",
              inputs: {
                command: "install"
              }
            },
            {
              task: "Gulp@0",
              displayName: "Gulp bundle",
              inputs: {
                gulpFile: "./gulpfile.js",
                targets: "bundle",
                arguments: "--ship"
              }
            },
            {
              task: "Gulp@0",
              displayName: "Gulp package",
              inputs: {
                targets: "package-solution",
                arguments: "--ship"
              }
            },
            {
              task: "Npm@1",
              displayName: "Install CLI for Microsoft 365",
              inputs: {
                command: "custom",
                verbose: false,
                customCommand: "install -g @pnp/cli-microsoft365"
              }
            },
            {
              script: "\n{{login}} \nm365 spo set --url '$(SharePointBaseUrl)' \n{{addApp}} \n{{deploy}}\n",
              displayName: "CLI for Microsoft 365 Deploy App"
            }
          ]
        }
      ]
    }
  ]
};