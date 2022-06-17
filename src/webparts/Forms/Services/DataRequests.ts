import { WebPartContext } from "@microsoft/sp-webpart-base";
import {SPHttpClient, ISPHttpClientOptions} from "@microsoft/sp-http";

const getListItems = async (context: WebPartContext, listUrl: string, listName: string, listDisplayName: string, pageSize: number, followedDocs: any) =>{
  
  const listData: any = [];
  const responseUrl = `${listUrl}/_api/web/Lists/GetByTitle('${listDisplayName}')/items?$top=${pageSize}&$select=id,Title,Created,DeptSubDeptGroupings,DeptSubDeptGroupings,FieldValuesAsText/FileRef&$expand=FieldValuesAsText`;
  
  try{
    const response = await context.spHttpClient.get(responseUrl, SPHttpClient.configurations.v1); //.then(r => r.json());

    if (response.ok){
      const results = await response.json();
      if(results){
        results.value.map((item: any)=>{
          listData.push({
            id: item.Id,
            title: item.Title || "",
            name: item.FieldValuesAsText.FileRef ? item.FieldValuesAsText.FileRef.substring(item.FieldValuesAsText.FileRef.lastIndexOf('/')+1, item.FieldValuesAsText.FileRef.lastIndexOf('.')) : "" ,
            link: item.FieldValuesAsText.FileRef,
            fileType: item.FieldValuesAsText.FileRef.substring(item.FieldValuesAsText.FileRef.lastIndexOf('.')+1),
            deptGrp: item.DeptSubDeptGroupings ? item.DeptSubDeptGroupings.substring(0, item.DeptSubDeptGroupings.indexOf('|')) : "",
            subDeptGrp: item.DeptSubDeptGroupings ? item.DeptSubDeptGroupings.substring(item.DeptSubDeptGroupings.indexOf('|')+1) : "",
            depts: item.DeptSubDeptGroupings ? item.DeptSubDeptGroupings : "",
            listUrl: listUrl,
            listName: listName,
            listDisplayName: listDisplayName,
            created: item.Created,
            details: "",
            webUrl: item.FieldValuesAsText['@odata.id'] ? item.FieldValuesAsText['@odata.id'].substring(item.FieldValuesAsText['@odata.id'].indexOf('.com')+5, item.FieldValuesAsText['@odata.id'].indexOf('/_api')) : "",
            listId: item['@odata.editLink'] ? item['@odata.editLink'].substring(item['@odata.editLink'].indexOf('guid')+5, item['@odata.editLink'].indexOf(')/Items')-1) : "",
            isFollowing: item.FieldValuesAsText.FileRef ? ((followedDocs.filter(driveItem => driveItem.name === item.FieldValuesAsText.FileRef.substring(item.FieldValuesAsText.FileRef.lastIndexOf('/')+1))).length > 0 ? true : false) : ''
          });
        });
      }
    }else{
      console.log("Forms response Error: " + listUrl + listName + response.statusText);
      return [];
    }
  }catch(error){
    console.log("Forms Error: " + error);
  }
  
  listData.sort((a,b) => a.name.localeCompare(b.name));
  
  return listData;
};

export const readAllLists = async (context: WebPartContext, listUrl: string, listName: string, pageSize: number, followedDocs: any) =>{
  const listData: any = [];
  let aggregatedListsPromises : any = [];
  const responseUrl = `${listUrl}/_api/web/Lists/GetByTitle('${listName}')/items`;

  try{
    const response = await context.spHttpClient.get(responseUrl, SPHttpClient.configurations.v1);

    if (response.ok){
      const responseResults = await response.json();
    
      responseResults.value.map((item: any)=>{
        listData.push({
          listName: item.Title,
          listDisplayName: item.ListDisplayName,
          listUrl: item.ListUrl
        });
      });

      listData.map((listItem: any)=>{
        aggregatedListsPromises = aggregatedListsPromises.concat(getListItems(context, listItem.listUrl, listItem.listName, listItem.listDisplayName, pageSize, followedDocs));
      });

    }else{
      console.log("Forms Error: " + listUrl + listName + response.statusText);
      return [];
    }
  }catch(error){
    console.log("Forms response error: " + error);
  }

  return Promise.all(aggregatedListsPromises);
};

export const getFollowed = async (context: WebPartContext) => {
  const graphResponse = await context.msGraphClientFactory.getClient();
  const followedDocsResponse = await graphResponse.api(`/me/drive/following`).top(1000).get();
  console.log("My Followed documents", followedDocsResponse);
  return followedDocsResponse;
};

const getDocDriveInfo = async (context: WebPartContext, listId: string, listItemId: string, webUrl: string) => {

  /** Steps
   * ok 1- listId -> get from list item
   * ok 2- listItemId -> get from list item
   * ok 3- siteUrl -> get from list item
   * ok 4- hostName -> pdsb1.sharepoint.com
   * ok 5- siteId -> 4529b386-b371-4b68-a258-23483541112b
   * 6- webId -> get from graph api 1st call -> https://graph.microsoft.com/v1.0/sites/<hostName>:/<webUrl>
   * 7- driveId & driveItemId-> get from graph api 2nd call -> https://graph.microsoft.com/v1.0/sites/<siteId,webId>/lists/<listId>/items/14/driveItem
   */

  const hostName = 'pdsb1.sharepoint.com';

  const graphResponse = await context.msGraphClientFactory.getClient();

  const webIdResponse = await graphResponse.api(`/sites/${hostName}:/${webUrl}`).get();
  const completeSiteId = webIdResponse.id;
  const siteId = completeSiteId.substring(completeSiteId.indexOf(',')+1, completeSiteId.lastIndexOf(','));
  const webId = completeSiteId.substring(completeSiteId.lastIndexOf(',')+1);

  const driveResponse = await graphResponse.api(`/sites/${siteId},${webId}/lists/${listId}/items/${listItemId}/driveItem`).get();

  console.log("webUrl", webUrl);
  console.log("completeSiteId", completeSiteId);
  console.log("siteId", siteId);
  console.log("webId", webId);
  console.log("driveResponse", driveResponse);


  return [driveResponse.parentReference.driveId, driveResponse.id];
};

export const unFollowDocument = async (context: WebPartContext, listId: string, listItemId: string, webUrl: string) => {
  const [driveId, driveItemId] =  await getDocDriveInfo(context, listId, listItemId, webUrl);

  const graphResponse = await context.msGraphClientFactory.getClient();
  const unfollowResponse = await graphResponse.api(`/drives/${driveId}/items/${driveItemId}/unfollow`).post(JSON.stringify(''));

  console.log("unfollowResponse", unfollowResponse);
};

export const followDocument = async (context: WebPartContext, listId: string, listItemId: string, webUrl: string) => {

  const [driveId, driveItemId] =  await getDocDriveInfo(context, listId, listItemId, webUrl);

  console.log("driveId", driveId);
  console.log("driveItemId", driveItemId);

  const graphResponse = await context.msGraphClientFactory.getClient();
  const followResponse = await graphResponse.api(`/drives/${driveId}/items/${driveItemId}/follow`).post(JSON.stringify(''));

  console.log("followResponse", followResponse);
};

export const isObjectEmpty = (items: any): boolean=>{
  let isEmpty:boolean = true;
  for (const item in items){
    isEmpty = items[item] === "" && isEmpty ;
  }
  return isEmpty;
};

export const arrayUnique = (arr, uniqueKey) => {
  const flagList = [];
  return arr.filter((item) => {
    if (flagList.indexOf(item[uniqueKey]) === -1) {
      flagList.push(item[uniqueKey]);
      return true;
    }
  });
};


