import { WebPartContext } from "@microsoft/sp-webpart-base";
import {SPHttpClient} from "@microsoft/sp-http";

const getListItems = async (context: WebPartContext, listUrl: string, listName: string, listDisplayName: string, pageSize: number) =>{
  
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
            depts: item.DeptSubDeptGroupings || "",
            listUrl: listUrl,
            listName: listName,
            listDisplayName: listDisplayName,
            created: item.Created,
            details: ""
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

export const readAllLists = async (context: WebPartContext, listUrl: string, listName: string, pageSize: number) =>{
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
        aggregatedListsPromises = aggregatedListsPromises.concat(getListItems(context, listItem.listUrl, listItem.listName, listItem.listDisplayName, pageSize));
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

export const isObjectEmpty = (items: any): boolean=>{
  let isEmpty:boolean = true;
  for (const item in items){
    isEmpty = items[item] === "" && isEmpty ;
  }
  return isEmpty;
};


export const arrayUnique = (arr, uniqueKey) => {
  const flagList = [];
  return arr.filter(function(item) {
    if (flagList.indexOf(item[uniqueKey]) === -1) {
      flagList.push(item[uniqueKey]);
      return true;
    }
  });
};

