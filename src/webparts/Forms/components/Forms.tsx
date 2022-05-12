import * as React from 'react';
import styles from './Forms.module.scss';
import {Icon, MessageBar, MessageBarType} from '@fluentui/react';
import { IFormsProps } from './IFormsProps';
import {readAllLists, getFollowed, unFollowDocument, followDocument} from  '../Services/DataRequests';
import IListItems from './IListItems/IListItems';
import IFilterFields from './IFilterFields/IFilterFields';
import toast, { Toaster } from 'react-hot-toast';

export default function MyTasks (props: IFormsProps){

  const [listItems, setListItems] = React.useState([]);
  const [preloaderVisible, setPreloaderVisible] = React.useState(true);
  const [filterFields, setFilterFields] = React.useState({
    name: "",
    depts: ""
  });

  const popToast = (toastMsg: string) =>{
    toast.success(toastMsg, {
      duration: 2000,
	  position: 'bottom-center',
      style: {
        margin: '20px',
		backgroundColor: '#616161',
		color: '#ffffff'
      },
    });
  };

  const fetchLists = () => {
	getFollowed(props.context).then(followedDocs => {
		readAllLists(props.context, props.listUrl, props.listName, props.pageSize, followedDocs.value).then((r: any) =>{
			setListItems(r.flat());
			setPreloaderVisible(false);
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


//   const followDocumentHandler = (item) => {
// 	setFollowPreloaderVisible(true);
// 	followDocument(props.context, item.listId, item.id, item.webUrl).then(() => {
// 		fetchLists();
// 	});
//   };
//   const unFollowDocumentHandler = (item) => {
// 	setFollowPreloaderVisible(true);
// 	unFollowDocument(props.context, item.listId, item.id, item.webUrl).then(()=>{
// 		fetchLists();
// 	});
//   };

  const followDocumentHandler = (item) => {
    followDocument(props.context, item.listId, item.id, item.webUrl).then(()=>{
		popToast('Added to Favorites!');
    });
	setListItems(prevState => {
		return prevState.map(prevItem => {
			const updatedItem = {...prevItem};
			if (updatedItem.listId === item.listId && updatedItem.id === item.id)
				updatedItem.isFollowing = true;
			return {...updatedItem};
		});
	});
  };
  const unFollowDocumentHandler = (item) => {
    unFollowDocument(props.context, item.listId, item.id, item.webUrl).then(()=>{
		popToast('Removed from Favorites!');
	});
	setListItems(prevState => {
		return prevState.map(prevItem => {
			const updatedItem = {...prevItem};
			if (updatedItem.listId === item.listId && updatedItem.id === item.id)
				updatedItem.isFollowing = false;
			return {...updatedItem};
		});
	});
  };
  
  return (
		<div className={styles.Forms}>
			<Toaster />

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
			/>
		</div>
  );
}
