import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { DisplayMode } from '@microsoft/sp-core-library';

import * as strings from 'AccordionWebPartStrings';
import { Accordion } from './components/Accordion';
import { IAccordionItem, IAccordionProps } from './components/IAccordionProps';

export interface IAccordionWebPartProps {
  items: IAccordionItem[];
  accordionMode: 'single' | 'openAll' | 'closeAll';
  headerLevel: string;
  fontSize: string;
}

export default class AccordionWebPart extends BaseClientSideWebPart<IAccordionWebPartProps> {

  public render(): void {
    const element: React.ReactElement<any> = React.createElement(
      Accordion,
      {
        items: this.properties.items || [],
        mode: this.properties.accordionMode || 'single',
        isEditMode: this.displayMode === DisplayMode.Edit,
        headerLevel: this.properties.headerLevel || 'h2',
        fontSize: this.properties.fontSize || '1rem',
        onUpdate: (items) => {
          this.properties.items = items;
          this.render();
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    if (!this.properties.items || this.properties.items.length === 0) {
      this.properties.items = [
        {
          id: '1',
          title: 'Accordion Section 1',
          content: '<p>Edit this content</p>',
          isOpen: false
        }
      ];
    }
    return super.onInit();
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          //header: { description: "Accordion Settings" },
          groups: [
            {
              groupName: "Accordion Behavior",
              groupFields: [
                PropertyPaneDropdown('accordionMode', {
                  label: 'Accordion Behavior',
                  options: [
                    { key: 'single', text: 'Open one at a time' },
                    { key: 'openAll', text: 'Open all' }
                  ]
                })
              ]
            },
            {
              groupName: "Header Formatting",
              groupFields: [
                PropertyPaneDropdown('headerLevel', {
                  label: 'Header Level',
                  options: [
                    { key: 'h1', text: 'Heading 1' },
                    { key: 'h2', text: 'Heading 2' },
                    { key: 'h3', text: 'Heading 3' },
                    { key: 'h4', text: 'Heading 4' },
                    { key: 'h5', text: 'Heading 5' },
                    { key: 'h6', text: 'Heading 6' }
                  ]
                }),
                PropertyPaneDropdown('fontSize', {
                  label: 'Font Size',
                  options: [
                    { key: '0.875rem', text: 'Small (14px)' },
                    { key: '1rem', text: 'Medium (16px)' },
                    { key: '1.125rem', text: 'Large (18px)' },
                    { key: '1.25rem', text: 'Extra Large (20px)' },
                    { key: '1.5rem', text: 'XXL (24px)' }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
