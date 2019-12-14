# Introduction to GitHub Actions

In our GitHub repositories, wouldn't it be nice to have a label on an issue automatically when it gets created or say when there is a pull request on the dev branch package and deploy the code to the dev site or say get a list of all the issues that are marked as P1 at the beginning of every week? That's where GitHub Actions come into picture. GitHub Actions can be thought of as code that runs when an event happens in GitHub repository.

To understand GitHub Actions, let's look into what GitHub workflows are. Consider a simple scenario - whenever a file pushed to a GitHub repo, we need to print "Hello world". That's a simple workflow. When we create such a workflow we write the code to print "Hello world" directly in the workflow. A workflow is written using yaml and it would look something like below:

```yaml
name: Print Hello world

on: push

jobs:
  greet:

    runs-on: windows-latest

    steps:
    - name: Run a one-line script
      run: Write-Host Hello world!
```

Now consider a slightly complex scenario - whenever a file is pushed to a GitHub repo, we need to print "Hello world" and then we need to build the code, deploy it and send an email to admin after deployment. We can choose write code for all these tasks (build, deploy and send mail) within the workflow itself. However, since these are common tasks for most of the projects, some developers would have written code for these already and hosted it in a different repositories i.e. there would be code for building the project, code for deploying and code for sending email. These are called GitHub Actions. 
So we what we can do is, consume that code (i.e. GitHub Actions) in our workflow. The structure of the workflow file (yaml) would look something like below:

```yaml
name: Print Hello world and Deploy

on: push

jobs:
  greet:

    runs-on: windows-latest

    steps:
    
    # Print Hello world
    - name: Run a one-line script
      run: Write-Host Hello world!
    
    # Step to Build the code, uses build action created by a community member
    - name: Build the code
      uses: build-action-created-by-a-community-member
    
    # Step to Deploy the code, uses deploy action created by a community member
    - name: Deploy the code
      uses: deploy-action-created-by-a-community-member
    
    # Step to Send an email, uses send-email action created by a community member
    - name: Send an email
      uses: send-email-action-created-by-a-community-member
```

Now that we have an understanding of workflow and GitHub Actions, we can start by [creating a simple workflow](./simpleworkflow.md).