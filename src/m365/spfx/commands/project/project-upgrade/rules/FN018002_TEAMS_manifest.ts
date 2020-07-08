import { Finding, Occurrence } from "../";
import { Project, Manifest } from "../../model";
import { Rule } from "./Rule";
import * as path from 'path';
import * as fs from 'fs';

export class FN018002_TEAMS_manifest extends Rule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN018002';
  }

  get title(): string {
    return 'Web part Microsoft Teams tab manifest';
  }

  get description(): string {
    return 'Create Microsoft Teams tab manifest for the web part';
  }

  get resolution(): string {
    return `add_cmd[BEFOREPATH]__filePath__[AFTERPATH][BEFORECONTENT]
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.2/MicrosoftTeams.schema.json",
  "manifestVersion": "1.2",
  "packageName": "__packageName__",
  "id": "__id__",
  "version": "0.1",
  "developer": {
    "name": "SPFx + Teams Dev",
    "websiteUrl": "https://products.office.com/en-us/sharepoint/collaboration",
    "privacyUrl": "https://privacy.microsoft.com/en-us/privacystatement",
    "termsOfUseUrl": "https://www.microsoft.com/en-us/servicesagreement"
  },
  "name": {
    "short": "__name__"
  },
  "description": {
    "short": "__description__",
    "full": "__description__"
  },
  "icons": {
    "outline": "tab20x20.png",
    "color": "tab96x96.png"
  },
  "accentColor": "#004578",
  "configurableTabs": [
    {
      "configurationUrl": "https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=__id__",
      "canUpdateConfiguration": true,
      "scopes": [
        "team"
      ]
    }
  ],
  "validDomains": [
    "*.login.microsoftonline.com",
    "*.sharepoint.com",
    "*.sharepoint-df.com",
    "spoppe-a.akamaihd.net",
    "spoprod-a.akamaihd.net",
    "resourceseng.blob.core.windows.net",
    "msft.spoppe.com"
  ],
  "webApplicationInfo": {
    "resource": "https://{teamSiteDomain}",
    "id": "00000003-0000-0ff1-ce00-000000000000"
  }
}
[AFTERCONTENT]`;
  };

  get resolutionType(): string {
    return 'cmd';
  }

  get file(): string {
    return '';
  };

  get severity(): string {
    return 'Optional';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.manifests ||
      project.manifests.length < 1) {
      return;
    }

    const webPartManifests: Manifest[] = project.manifests.filter(m => m.componentType === 'WebPart');
    if (webPartManifests.length < 1) {
      return;
    }

    const occurrences: Occurrence[] = [];
    webPartManifests.forEach(manifest => {
      const webPartFolderName: string = path.basename(path.dirname(manifest.path));
      const teamsFolderName: string = `teams`;
      const teamsFolderPath: string = path.join(project.path, teamsFolderName);
      const teamsManifestPath: string = path.join(teamsFolderPath, `manifest_${webPartFolderName}.json`);
      if (fs.existsSync(teamsManifestPath)) {
        return;
      }

      let webPartTitle: string = 'undefined';
      let webPartDescription: string = 'undefined';
      if (manifest.preconfiguredEntries &&
        manifest.preconfiguredEntries.length > 0) {
        const entry = manifest.preconfiguredEntries[0];
        if (entry.title && entry.title.default) {
          webPartTitle = entry.title.default;
        }
        if (entry.description && entry.description.default) {
          webPartDescription = entry.description.default;
        }
      }
      const webPartId: string = manifest.id || 'undefined';

      const resolution: string = this.resolution
        .replace(/__filePath__/g, teamsManifestPath)
        .replace(/__packageName__/g, webPartTitle)
        .replace(/__id__/g, webPartId)
        .replace(/__name__/g, webPartTitle)
        .replace(/__description__/g, webPartDescription);

      occurrences.push({
        file: path.relative(project.path, teamsManifestPath),
        resolution: resolution
      });
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}
