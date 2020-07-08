import { Finding, Occurrence } from "../";
import { Project } from "../../model";
import { TsRule } from "./TsRule";
import * as ts from 'typescript';
import * as os from 'os';

export class FN016004_TS_property_pane_property_import extends TsRule {
  constructor() {
    super();
  }

  get id(): string {
    return 'FN016004';
  }

  get title(): string {
    return 'Property pane property import change to @microsoft/sp-property-pane';
  }

  get description(): string {
    return `Refactor the code to import property pane property from the @microsoft/sp-property-pane npm package instead of the @microsoft/sp-webpart-base package`;
  }

  get resolution(): string {
    return '';
  };

  get resolutionType(): string {
    return 'ts';
  }

  get severity(): string {
    return 'Required';
  }

  visit(project: Project, findings: Finding[]): void {
    if (!project.tsFiles) {
      return;
    }

    const propertyPaneObjects: string[] = ['PropertyPaneLifeCycleEvent', 'ILifeCycleEventCallback', '_PropertyPaneController', 'PropertyPaneAction', 'IPropertyPaneData', 'IPropertyPaneConfiguration', 'IPropertyPanePage', 'IPropertyPanePageHeader', 'IPropertyPaneGroup', 'IPropertyPaneField', 'PropertyPaneFieldType', 'IPropertyPaneCustomFieldProps', 'PropertyPaneCustomField', 'IPropertyPaneButtonProps', 'PropertyPaneButtonType', 'PropertyPaneButton', 'IPropertyPaneCheckboxProps', 'PropertyPaneCheckbox', 'IPropertyPaneChoiceGroupOptionIconProps', 'IPropertyPaneChoiceGroupProps', 'IPropertyPaneChoiceGroupOption', '_IPropertyPaneChoiceGroupOptionInternal', 'PropertyPaneChoiceGroup', 'IPropertyPaneDropdownProps', 'IPropertyPaneDropdownOption', 'PropertyPaneDropdownOptionType', 'IPropertyPaneDropdownCalloutProps', 'PropertyPaneDropdown', 'IPropertyPaneDynamicFieldFilters', 'IPropertyPaneDynamicFieldProps', 'PropertyPaneDynamicField', 'PropertyPaneDynamicFieldSet', 'IPropertyPaneDynamicFieldSetProps', 'PropertyPaneHorizontalRule', 'IPropertyPaneLabelProps', 'PropertyPaneLabel', 'IPropertyPaneLinkProps', 'PropertyPaneLink', 'IPropertyPaneSliderProps', 'PropertyPaneSlider', 'IPropertyPaneTextFieldProps', 'PropertyPaneTextField', 'IPropertyPaneDynamicTextFieldProps', 'IConfiguredDynamicTextFieldProps', 'PropertyPaneDynamicTextField', 'IPropertyPaneToggleProps', 'PropertyPaneToggle', 'IPropertyPaneSpinButtonProps', 'PropertyPaneSpinButton', 'IPropertyPaneConsumer', 'IDynamicConfiguration', '_IDynamicConfiguration', 'IDynamicDataSharedSourceConfiguration', 'IDynamicDataSharedPropertyConfiguration', 'DynamicDataSharedDepth', 'IDynamicDataSharedSourceFilters', 'IDynamicDataSharedPropertyFilters', 'IPropertyPaneConditionalGroup'];

    const occurrences: Occurrence[] = [];
    project.tsFiles.forEach(file => {
      const nodes: ts.Node[] | undefined = file.nodes;
      if (!nodes) {
        return;
      }

      const obj: ts.ImportDeclaration[] = nodes
        .filter(n => ts.isImportDeclaration(n))
        .map(n => n as ts.ImportDeclaration)
        .filter(n => n.getText().indexOf('@microsoft/sp-webpart-base') > 0);

      obj.forEach(n => {
        const resource: string = n.getText();
        const importsText: string = resource.replace(/\s/g, '').substr(resource.indexOf('{'));
        const imports: string[] = importsText.substr(0, importsText.indexOf('}')).split(',');
        const importsToStay: string[] = [];
        const importsToBeMoved: string[] = [];

        imports.forEach(importName => {
          if (propertyPaneObjects.indexOf(importName) > -1) {
            importsToBeMoved.push(importName);
          }
          else {
            importsToStay.push(importName);
          }
        })

        if (importsToBeMoved.length > 0) {
          const newBaseImportDeclaration: string = `import { ${importsToStay.join(', ')} } from "@microsoft/sp-webpart-base";`;
          const newPropertiesImportDeclaration: string = `import { ${importsToBeMoved.join(', ')} } from "@microsoft/sp-property-pane";`;
          let resolution: string = newPropertiesImportDeclaration;
          if (importsToStay.length > 0) {
            resolution = `${newBaseImportDeclaration}${os.EOL}${resolution}`;
          }

          this.addOccurrence(resolution, file.path, project.path, n, occurrences);
        }
      });
    });

    if (occurrences.length > 0) {
      this.addFindingWithOccurrences(occurrences, findings);
    }
  }
}