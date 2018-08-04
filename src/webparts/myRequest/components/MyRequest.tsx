import * as React from 'react';
import styles from './MyRequest.module.scss';
import { IMyRequestProps } from './IMyRequestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Image, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { DetailsList, buildColumns, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { IPersonaSharedProps, Persona, PersonaSize, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { TaxonomyPicker } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import { IPickerTerms } from '@pnp/spfx-controls-react/lib/controls/taxonomyPicker';
var classNames = require('classnames');

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export interface IMyRequestState {
  items?: any[];
  columns?: IColumn[];
}

export default class MyRequest extends React.Component<IMyRequestProps, IMyRequestState> {
  private data;
  constructor(props) {
    super(props);
    this.reloaData();
  }

  private reloaData = () => {
    console.log('reloaData');
    this.data = this.createListItems(10);
    this.state = {
      items: this.data,
      columns: this._buildColumns()
    };
    this.render();
  };

  public render(): React.ReactElement<IMyRequestProps> {
    const { items, columns } = this.state;
    return (
      <div className={styles.myRequest}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div>
              <span className={styles.title}>My requests</span>

              <div className={styles.row} style={{ 'width': '300px' }}>
                <TaxonomyPicker
                  allowMultipleSelections={false}
                  termsetNameOrID='Offices'
                  panelTitle="Select an office"
                  label=""
                  context={this.props.context}
                  onChange={this.onPickerChange}
                  isTermSetSelectable={false}
                />
              </div>
              <div className={styles.row}>
                <DetailsList
                  items={items as any[]}
                  setKey="set"
                  columns={columns}
                  onRenderItemColumn={this._renderItemColumn}
                  onColumnHeaderClick={this._onColumnClick}
                  onItemInvoked={this._onItemInvoked}
                  onColumnHeaderContextMenu={this._onColumnHeaderContextMenu}
                />

              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private onPickerChange = (terms: IPickerTerms) => {
    this.reloaData();
    
  }
  private _onColumnClick = (event: React.MouseEvent<HTMLElement>, column: IColumn): void => {

  };

  private _onColumnHeaderContextMenu(column: IColumn | undefined, ev: React.MouseEvent<HTMLElement> | undefined): void {

  }

  private _onItemInvoked(item: any, index: number | undefined): void {
  }

  private _renderItemColumn = (item: any, index: number, column: IColumn) => {

    const fieldContent = item[column.fieldName || ''];
    switch (column.key) {
      case 'FILENAME':
        return (<span data-selection-disabled={true}>
          <FileTypeIcon type={IconType.image} path={"https://contoso.sharepoint.com/documents/" + fieldContent} />
          <span style={{ "marginLeft": "10px" }}>{fieldContent}</span>
        </span>);

      case 'REQUESTTO':
        return (<span data-selection-disabled={true}>
          <Persona
            {...fieldContent}
            size={PersonaSize.size24}
            presence={fieldContent.presence}
            hidePersonaDetails={false}
          />
        </span>);

      case 'PROGRESS':
        var classes = 'ms-bgColor-red';
        if (parseInt(fieldContent) == 3) {
          classes = 'ms-bgColor-yellow';
        }
        if (parseInt(fieldContent) > 3) {
          classes = 'ms-bgColor-tealLight';
        }

        var width = (parseInt(fieldContent) / 5) * 100;


        return <span>
          <div>
            {fieldContent}/5
          </div>
          <div className={'ms-bgColor-neutralLight'}>
            <div style={{
              "width": width.toString() + "%",
              "height": "5px",
              "display": "block"
            }}
              className={'ms-fontColor-black sp-field-dataBars ' + classes}>
            </div></div>
        </span>;

      default:
        return <span>{fieldContent}</span>;
    }
  }

  private _buildColumns = () => {
    const columns = buildColumns(this.data);

    const fileNameCol = columns.filter(column => column.name === 'FILENAME')[0];
    // Special case one column's definition.
    fileNameCol.name = 'FILE NAME';
    fileNameCol.minWidth = 200;

    const requestToCol = columns.filter(column => column.name === 'REQUESTTO')[0];
    // Special case one column's definition.
    requestToCol.name = 'REQUEST TO';
    requestToCol.minWidth = 200;
    requestToCol.isResizable = true;


    const pogressToCol = columns.filter(column => column.name === 'PROGRESS')[0];
    pogressToCol.minWidth = 200;
    return columns;
  }



  private DATA = {
    created: ['Apr 12, 2018', 'Mar 29, 2017', 'May 15, 2018', 'Jan 12, 2018'],
    workflows: ['Payment', 'Promotion', 'Assign budget'],
    files: ['Prototype.docx'
      , 'Payslip.xls'
      , 'BSR-FDS.onetoc'
      , 'Ericsson_Change_Request.pptx'],
    progress: [1, 2, 3, 4, 5],
    people: [{
      imageUrl: 'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-female.png',
      imageInitials: 'CH',
      primaryText: 'Chau Huynh',
      secondaryText: 'Senior Developer',
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
        FILENAME: this.randomItem(this.DATA.files),
        CREATED: this.randomItem(this.DATA.created),
        WORKFLOW: this.randomItem(this.DATA.workflows),
        REQUESTTO: this.randomItem(this.DATA.people),
        PROGRESS: this.randomItem(this.DATA.progress),
      };
    });
  }


  private randomItem = (array: any[]): any => {
    const index = Math.floor(Math.random() * array.length);
    return array[index];
  }
}
