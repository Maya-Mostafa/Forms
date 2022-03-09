import * as React from 'react';
import {IListItemsProps} from './IListItemsProps';
import styles from '../Forms.module.scss';
import {MessageBar, MessageBarType, Spinner} from '@fluentui/react';
import { ListView, IViewField, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { TooltipHost, Icon } from '@fluentui/react';
import { initializeFileTypeIcons } from '@uifabric/file-type-icons';
import { getFileTypeIconProps } from '@uifabric/file-type-icons';

// Register icons and pull the fonts from the default SharePoint cdn.
initializeFileTypeIcons();

export interface IDocument {
    key: string;
    name: string;
    value: string;
    iconName: string;
    fileType: string;
    modifiedBy: string;
    dateModified: string;
    dateModifiedValue: number;
    fileSize: string;
    fileSizeRaw: number;
  }

export default function IListItems (props: IListItemsProps) {
  
  const viewFields:IViewField [] = [
    {
        name: '',
        displayName: '',
        minWidth: 16,
        maxWidth: 16,
        render: (item: IDocument) => (
          <TooltipHost content={`${item.fileType} file`}>
            <Icon {...getFileTypeIconProps({extension: item.fileType, size: 16}) }/>
          </TooltipHost>
        ),
    },
    {
        name: 'name',
        displayName : 'Form Name',
        minWidth: 150,
        maxWidth: 450,
        sorting: true,
        isResizable: true,
        render : (item: any) => (
        <div>
            <a className={styles.defautlLink} target="_blank" data-interception="off" href={item.link}>{item.name}</a>
        </div>
        )
    },
  ];
  const groupByFields: IGrouping[] = [
    {
        name: "deptGrp", 
        order: GroupOrder.ascending 
    },
    {
        name: "subDeptGrp", 
        order: GroupOrder.ascending 
    }
  ];

  const filteredItems = (props.items.filter((listItem: any)=>{
    let filterFieldVal: string;
    for (let i in props.filterField) {
        filterFieldVal = props.filterField[i];
        if (listItem[i] === undefined || listItem[i].toString().toLowerCase().indexOf(filterFieldVal.toLowerCase()) === -1)
            return false;
    }
    return true;
  }));
  

  return(
    <div className={styles.listViewNoWrap}>
        <ListView
            items={filteredItems}
            viewFields={viewFields}
            groupByFields={groupByFields}
            // stickyHeader={true} 
            compact={true}
        />
        {filteredItems.length === 0 && !props.preloaderVisible &&
            <MessageBar
                messageBarType={MessageBarType.warning}
                isMultiline={false}>
                Sorry, there is no data to display.
            </MessageBar>
        } 
        {props.preloaderVisible &&
            <div>
                <Spinner label="Loading data, please wait..." ariaLive="assertive" labelPosition="right" />
            </div>
        }
    </div>
  );
}





