#!/bin/bash
set -e
set -o pipefail

if [[ -z "$APP_FILE_PATH" ]]; then
  echo "App file path - APP_FILE_PATH - not set."
  exit 1
fi

if [[ ! -d "$APP_FILE_PATH" &&  ! -f "$APP_FILE_PATH" ]];
then
    echo "Invalid file path '$APP_FILE_PATH'."
    exit 1
fi

if [[ "$SCOPE" == 'sitecollection' ]]; then
  if [[ -z "$SITE_COLLECTION_URL" ]]; then
    echo "Site collection URL - SITE_COLLECTION_URL - is needed when scope is set to sitecollection."
    exit 1
  fi
fi

main() {
    
    echo "Starting upload and deployment..."
    if [[ "$SCOPE" == 'sitecollection' ]]; then
      appId=$(o365 spo app add -p $APP_FILE_PATH --overwrite --scope sitecollection --appCatalogUrl $SITE_COLLECTION_URL)
      o365 spo app deploy --name $(basename $APP_FILE_PATH) --scope sitecollection --appCatalogUrl $SITE_COLLECTION_URL
      o365 spo app install --id $appId --siteUrl $SITE_COLLECTION_URL --scope sitecollection
    else
      o365 spo app add -p $APP_FILE_PATH --overwrite
      o365 spo app deploy --name $(basename $APP_FILE_PATH)
    fi
    echo "Upload and deployment complete."
}

main "$@"