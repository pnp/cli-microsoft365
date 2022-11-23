# What to expect during a Pull request review

At CLI for Microsoft 365, we love it when you send us a PR (pull request) for an issue that you are helping us with. The review process here is a collaborative effort between you and the maintainers who review your code. You might see one or more than one reviewer working on your PR based on availability.  

At CLI for Microsoft 365, we strive to process PRs as quickly as possible. We process PRs in the order in which they arrive (first in first out). PRs with bug fixes take priority over new features and enhancements because they address active issues that block our users. When a PR was reviewed and needs adjustments, it goes to the end of the processing queue so that it's not blocking the review process. 

This article will help you to understand further how we review a pull request, what you can expect during the review process and how a pull request will be merged or closed. 

## What happens immediately after you send us a PR? 

One of our maintainers will have a look into the PR and will check if the PR build has passed before they can take up and review the PR. The automated checks include: 

- Does the code build 
- Do the automated tests pass and have 100% test coverage? 
- Our linter does not catch any issues in the source code (to ensure you follow certain naming conventions etc, to keep the source code consistent)
 
The oldest PRs are reviewed first and once reviewed will be moved to the end of the queue. In this case, the maintainers immediately alert the author to make sure they are rectified before they can review the PR. 

## What happens when someone reviews your PR? 

When your PR is reviewed by one of the maintainers, they will first assign themselves to the PR. This is a clear indication that it is getting reviewed. Here is the checklist with which the review is normally done by the maintainers. 

[PR checklist · pnp/cli-microsoft365 Wiki (github.com)](https://github.com/pnp/cli-microsoft365/wiki/PR-checklist)

They may also suggest some best practices to ensure a consistent coding style. If there are comments on the review, for code changes then you will see `changes requested` in your PR visible next to your username. If there are any questions related to the issue that the maintainer needs you to answer before continuing with the review, the PR will be labeled `waiting on response`. The reviewer will also then mark the pull request as **Draft**. This is for you to go ahead and make changes in the code as suggested and the reviewers know that this is still a work in progress

## What should you do after you have made changes? 

Once you have updated the code and you feel the PR is ready to be sent back to the maintainers for review, you can mark the PR as **Ready for review**. This will put the PR back into the backlog for the reviewers.  

## What happens when the PR is reviewed and ready? 

The reviewer will do another round of review and if some things still need to be changed, we repeat the cycle of a review or if everything is okay, we approve the PR. The reviewer will label it as `pr-merged`. This is the indication that your PR is completely approved and merged into the source code's main branch. You will be alerted with a comment from the reviewer as well that it is merged. The reviewer will then close the PR which in turn closes the issue as well. 

## How long will reviewers wait for your response once changes are requested? 

On a PR the reviewers will wait two weeks before they close the PR. They will do a follow-up with you before closing. You can open a new PR at any time after syncing your fork. Here is a [guide](https://github.com/pnp/cli-microsoft365/blob/main/CONTRIBUTING.md#tips) to help you sync your fork.
