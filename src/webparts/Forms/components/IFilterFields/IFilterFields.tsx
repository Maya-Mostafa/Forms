import * as React from 'react';
import {IFilterFieldsProps} from './IFilterFieldsProps';
import {Stack, IStackProps, IStackStyles, SearchBox, ActionButton, initializeIcons, ComboBox, IComboBoxOption, Icon} from '@fluentui/react';
import styles from '../Forms.module.scss';
import {isObjectEmpty} from '../../Services/DataRequests';

export default function IFilterFields (props: IFilterFieldsProps) {
    
    initializeIcons();
    const stackTokens = { childrenGap: 50 };
    const stackStyles: Partial<IStackStyles> = { root: { width: '100%' } };
    const columnProps: Partial<IStackProps> = {
        tokens: { childrenGap: 15 },
        styles: { root: { width: '50%' } },
    };
    
    return(
        <div className={styles.filterForm}>            
            <ActionButton 
                className={styles.resetSrchBtn}
                text="Reset" 
                onClick={props.resetSrch} 
                iconProps={{ iconName: 'ClearFilter' }}
                allowDisabledFocus 
                disabled = {isObjectEmpty(props.filterField)}
            />
            <div>
                <Stack horizontal tokens={stackTokens} styles={stackStyles}>
                    <Stack {...columnProps}>                        
                        <SearchBox 
                            placeholder="Form Name" 
                            value={props.filterField.name}
                            onChange={props.onChangeFilterField("name")}
                            iconProps={{ iconName: 'Rename' }}
                            showIcon={true}
                            underlined
                            className={styles.srchBox}
                        />
                        <SearchBox 
                            placeholder="Location" 
                            value={props.filterField.depts}
                            onChange={props.onChangeFilterField("depts")}
                            iconProps={{ iconName: 'GlobalNavButton' }}
                            showIcon={true}
                            underlined
                            className={styles.srchBox}
                        />
                    </Stack>
                    
                </Stack>
            </div>
        </div>
    );
}