# Forms Dashboard
- Displaying forms from different lists 
- links to forms
- forms grouping by departments and sub-departments
- search by form name and location
- favorite/unfavorite feature for forms
- user can view favorites by clicking on the favorites link that appears as a notification or by going to their office page

# Requests
- Rest calls for links list
- Graph API for follow/unfollow docs

# REST API Requests
- linksResponseUrl = `${listUrl}/_api/web/Lists/GetByTitle('${listName}')/items`;
- listResponseUrl = `${listUrl}/_api/web/Lists/GetByTitle('${listDisplayName}')/items?$top=${pageSize}&$select=id,Title,Created,DeptSubDeptGroupings,DeptSubDeptGroupings,FieldValuesAsText/FileRef&$expand=FieldValuesAsText`;

# Graph API Requests
- Get followed items https://graph.microsoft.com/v1.0/me/drive/following
- Unfollow drive item -> POST /drives/{drive-id}/items/{item-id}/unfollow
- follow drive item -> POST /drives/{drive-id}/items/{item-id}/follow

# Graph API Requests Steps
1- listId -> get from list item
2- listItemId -> get from list item
3- siteUrl -> get from list item
4- hostName -> pdsb1.sharepoint.com
5- siteId -> 4529b386-b371-4b68-a258-23483541112b
6- webId -> get from graph api 1st call -> https://graph.microsoft.com/v1.0/sites/<hostName>:/<webUrl>
7- driveId & driveItemId-> get from graph api 2nd call -> https://graph.microsoft.com/v1.0/sites/<siteId,webId>/lists/<listId>/items/14/driveItem