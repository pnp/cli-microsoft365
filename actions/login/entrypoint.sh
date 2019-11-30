#!/bin/bash
set -e
set -o pipefail

if [[ -z "$ADMIN_USERNAME" ]]; then
  echo "Admin user name - ADMIN_USERNAME - not set."
  exit 1
fi

if [[ -z "$ADMIN_PASSWORD" ]]; then
  echo "Admin password - ADMIN_PASSWORD - not set."
  exit 1
fi


main() {
    echo "Logging into tenant using O365 CLI..."
    o365 login --authType password --userName $ADMIN_USERNAME --password $ADMIN_PASSWORD
    o365 status  
    echo "Logged in."
}

main "$@"