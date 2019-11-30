#!/bin/bash

set -e

if [ -n "$O365_CLI_SCRIPT_PATH" ]
then
  SCRIPT_FILE="${GITHUB_WORKSPACE}/${O365_CLI_SCRIPT_PATH}"
  if [[ -e "$SCRIPT_FILE" ]]
  then
    chmod +x "$SCRIPT_FILE"
    $SCRIPT_FILE
  else
    echo "Script file ${SCRIPT_FILE} does not exists." >&2
    exit 1
  fi
else
  sh -c "$O365_CLI_SCRIPT"
fi