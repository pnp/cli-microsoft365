import { gitHubWorkflow } from "./project-github-workflow-model";

export const workflow: gitHubWorkflow = {
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
      steps: [
        {
          name: "Checkout",
          uses: "actions/checkout@v3.5.3"
        },
        {
          name: "Use Node.js 16.x",
          uses: "actions/setup-node@v3.7.0",
          with: {
            "node-version": "16.x"
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
          uses: "pnp/action-cli-login@v2.2.2",
          with: {
            "CERTIFICATE_ENCODED": "${{ secrets.CERTIFICATE_ENCODED }}",
            "CERTIFICATE_PASSWORD": "${{ secrets.CERTIFICATE_PASSWORD }}",
            "APP_ID": "${{ secrets.APP_ID }}"
          }
        },
        {
          name: "CLI for Microsoft 365 Deploy App",
          uses: "pnp/action-cli-deploy@v3.0.1",
          with: {
            "APP_FILE_PATH": "sharepoint/solution/{{ solutionName }}.sppkg",
            "SKIP_FEATURE_DEPLOYMENT": false,
            "OVERWRITE": false
          }
        }
      ]
    }
  }
};