import * as React from 'react';
import styles from './Forms.module.scss';
import {Icon, MessageBar, MessageBarType} from '@fluentui/react';
import { IFormsProps } from './IFormsProps';
import {readAllLists, getFollowed, unFollowDocument, followDocument} from  '../Services/DataRequests';
import IListItems from './IListItems/IListItems';
import IFilterFields from './IFilterFields/IFilterFields';

export default function MyTasks (props: IFormsProps){

  const [listItems, setListItems] = React.useState([]);
  const [preloaderVisible, setPreloaderVisible] = React.useState(true);
  const [isFollowPreloaderVisible, setFollowPreloaderVisible] = React.useState(false);
  const [filterFields, setFilterFields] = React.useState({
    name: "",
    depts: ""
  });

  const fetchLists = () => {
	getFollowed(props.context).then(followedDocs => {
		readAllLists(props.context, props.listUrl, props.listName, props.pageSize, followedDocs.value).then((r: any) =>{
			setListItems(r.flat());
			setPreloaderVisible(false);
			setFollowPreloaderVisible(false);
		});
	});
  };

  React.useEffect(()=>{
	fetchLists();
  }, []);

  const onChangeFilterField = (fieldNameParam: string) =>{
    return(ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: any) =>{   
      setFilterFields({
        ...filterFields,
        [fieldNameParam] : text || ""
      });
    };
  };
  
  const resetSrch = () =>{    
    setFilterFields({
      name: "",
      depts: ""
    });
  };


  const followDocumentHandler = (item) => {
	setFollowPreloaderVisible(true);
	followDocument(props.context, item.listId, item.id, item.webUrl).then(() => {
		fetchLists();
	});
  };
  const unFollowDocumentHandler = (item) => {
	setFollowPreloaderVisible(true);
	unFollowDocument(props.context, item.listId, item.id, item.webUrl).then(()=>{
		fetchLists();
	});
  };


  return (
		<div className={styles.Forms}>
			<h2>{props.wpTitle}</h2>

			<div className={styles.fieldsAndHelp}>
				<div className={styles.fieldsSection}>
					<IFilterFields
						filterField={filterFields}
						onChangeFilterField={onChangeFilterField}
						resetSrch={resetSrch}
					/>
				</div>
				{props.showHelp && (
					<div className={styles.helpSection}>
						<a
							href={props.helpLink}
							title={props.helpTitle}
							target='_blank'
							data-interception='off'
						>
							<Icon iconName='StatusCircleQuestionMark' />
						</a>
					</div>
				)}
			</div>

			{props.showHelpMsg && (
				<MessageBar
					messageBarType={MessageBarType.warning}
					isMultiline={true}
					className={styles.helpMsg}
				>
					{props.helpMsgTxt}
					<a href={props.helpMsgLink}>{props.helpMsgLinkTxt}</a>
				</MessageBar>
			)}

			<IListItems
				items={listItems}
				preloaderVisible={preloaderVisible}
				filterField={filterFields}
				followDocument={followDocumentHandler}
				unFollowDocument={unFollowDocumentHandler}
				isFollowPreloaderVisible={isFollowPreloaderVisible}
			/>
		</div>
  );
}
