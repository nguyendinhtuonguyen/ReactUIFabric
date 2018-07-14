import * as React from 'react';
import styles from './MyRequest.module.scss';
import { IMyRequestProps } from './IMyRequestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Image, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { DetailsList, buildColumns, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { IPersonaSharedProps, Persona, PersonaSize, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
var classNames = require('classnames');
export interface IMyRequestState {
  items?: any[];
  columns?: IColumn[];
}

export default class MyRequest extends React.Component<IMyRequestProps, IMyRequestState> {

  constructor(props) {
    super(props);

    this.state = {
      items: this.props.items,
      columns: this._buildColumns()
    };
  }

  public render(): React.ReactElement<IMyRequestProps> {
    const { items, columns } = this.state;
    return (
      <div className={styles.myRequest}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div>
              <span className={styles.title}>My requests</span>

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
        return( <span data-selection-disabled={true}>
          <FileTypeIcon type={IconType.image} path={"https://contoso.sharepoint.com/documents/" + fieldContent} />
          <span style={{ "marginLeft": "10px" }}>{fieldContent}</span>
        </span>);

      case 'REQUESTTO':
        return (<span data-selection-disabled={true}>
          <Persona
            {...fieldContent}
            size={PersonaSize.size24}
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
    const columns = buildColumns(this.props.items);
    
    const fileNameCol = columns.filter(column => column.name === 'FILENAME')[0];
    // Special case one column's definition.
    fileNameCol.name = 'FILE NAME';
    fileNameCol.maxWidth = 350;

    const requestToCol = columns.filter(column => column.name === 'REQUESTTO')[0];
    // Special case one column's definition.
    requestToCol.name = 'REQUEST TO';
    requestToCol.maxWidth = 200;

    return columns;
  }
}
