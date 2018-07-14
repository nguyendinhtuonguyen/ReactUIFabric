import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'MyRequestWebPartStrings';
import MyRequest from './components/MyRequest';
import { IMyRequestProps } from './components/IMyRequestProps';
import { IPersonaSharedProps, Persona, PersonaSize, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';

export interface IMyRequestWebPartProps {
  description: string;
}

export default class MyRequestWebPart extends BaseClientSideWebPart<IMyRequestWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMyRequestProps> = React.createElement(
      MyRequest,
      {
        description: this.properties.description,
        items: this.createListItems(10)
      }
    );

    ReactDom.render(element, this.domElement);
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
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }


  private LOREM_IPSUM = (
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut ' +
    'labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut ' +
    'aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore ' +
    'eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt ' +
    'mollit anim id est laborum'
  ).split(' ');

  private DATA = {
    created: ['Apr 12, 2018 by me', 'Mar 29, 2017 by Par', 'May 15, 2018 by Chau', 'Jan 12, 2018 by Toan'],
    workflows: ['Payment', 'Promotion', 'Assign budget'],
    files: ['Prototype.docx'
      , 'Payslip.xls'
      , 'BSR-FDS.onetoc'
      , 'Ericsson_Change_Request.pptx'],
    progress:[1,2,3,4,5],
    people: [{
      imageUrl: 'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-female.png',
      imageInitials: 'CH',
      primaryText: 'Chau Huynh',
      secondaryText: 'Adult Developer',
      tertiaryText: 'Online',
      showSecondaryText: true,
      presence: PersonaPresence.online
    },
    {
      imageUrl: 'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-male.png',
      imageInitials: 'PJ',
      primaryText: 'Pär Johansson',
      secondaryText: 'Director',
      tertiaryText: 'In a meeting',
      optionalText: 'Available at 4:00pm',
      showSecondaryText: true,
      presence: PersonaPresence.busy
    },
    {
      imageInitials: 'TD',
      primaryText: 'Toan Dinh',
      secondaryText: 'Developer',
      tertiaryText: 'Away',
      showSecondaryText: true,
      presence: PersonaPresence.away
    }]
  };
  
  private createListItems = (count: number, startIndex: number = 0): any => {
    return Array.apply(null, Array(count)).map((item: number, index: number) => {

      return {
        FILENAME: this.randWord(this.DATA.files),
        CREATED: this.randWord(this.DATA.created),
        WORKFLOW: this.randWord(this.DATA.workflows),
        REQUESTTO: this.randomItem(this.DATA.people),
        PROGRESS: this.randomItem(this.DATA.progress),
      };
    });
  }

  private lorem = (wordCount: number): string => {
    return Array.apply(null, Array(wordCount))
      .map((item: number) => this.randWord(this.LOREM_IPSUM))
      .join(' ');
  }

  private isGroupable = (key: string): boolean => {
    return key === 'color' || key === 'shape' || key === 'location';
  }

  private randWord = (array: string[]): string => {
    const index = Math.floor(Math.random() * array.length);
    return array[index];
  }

  private randomItem = (array: any[]): any => {
    const index = Math.floor(Math.random() * array.length);
    return array[index];
  }
}
