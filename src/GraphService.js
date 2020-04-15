import moment from 'moment';
var graph = require('@microsoft/microsoft-graph-client');

// https://docs.microsoft.com/en-us/graph/overview
// https://docs.microsoft.com/en-us/graph/query-parameters
// https://www.npmjs.com/package/msgraph-sdk-javascript

// function formatDateTime(dateTime) {
//   return moment.utc(dateTime).local().format('M/D/YY h:mm A');
// }
function getTodayDate() {
  return moment().format(); //Format: "2020-04-02T11:36:57-03:00"
}

function getTomorrowDate() {
  return moment().add(1, 'd').format('YYYY-MM-DDT00:00'); //Get start of next day 0:0 hours
}

function getAuthenticatedClient(accessToken) {
  // Initialize Graph client
  const client = graph.Client.init({
    // Use the provided access token to authenticate
    // requests
    authProvider: (done) => {
      done(null, accessToken.accessToken);
    }
  });

  return client;
}

export async function getUserDetails(accessToken) {
  const client = getAuthenticatedClient(accessToken);

  const user = await client.api('/me').get();
  return user;
}

export async function getEvents(accessToken) {
  const client = getAuthenticatedClient(accessToken);

  const events = await client
    .api('/me/events')
    // .select('subject,organizer,start,end')
    .filter(`start/dateTime gt '${getTodayDate()}' and start/dateTime lt '${getTomorrowDate()}'`)
    // .filter(`start/dateTime gt '3/27/20' and start/dateTime lt '3/28/20'`)
    // .filter(`start/dateTime gt '3/27/20'`)
    .orderby('start/dateTime ASC')
    .get();

  return events;
}

export async function getMessages(accessToken) {
  const client = getAuthenticatedClient(accessToken);

  const events = await client
    // .api('/me/messages')
    // .api('/me/mailFolders/inbox/messages')
    .api('/me/mailFolders/inbox/messages')
    // .select('subject,bodyPreview,body,sender') //TODO dsabled for testing
    .orderby('receivedDateTime DESC')
    .get();

  for (const msg of events.value) {
    let msgContent = msg.body.content;
    msgContent = msgContent.replace(/<img/g, '<img style="max-width: 100%;object-fit: scale-down;"' );
    let msgId = msg.id;
    let attachments = await client
      .api(`/me/messages/${msgId}/attachments`)
      .get();
    for (const a of attachments.value) {
      msgContent = msgContent.replace(`cid:${a.contentId}`, `data:image/jpeg;base64,${a.contentBytes}`);
    }

    //TODO add style to all images to prevent huge images
    //style="max-width: 100%;"

    msg.body.content = msgContent;
  };
  // let msgContent = events.value[0].body.content;
  // let msgId = events.value[0].id;
  // let cidList = msgContent.match(/cid["']?((?:.(?!["']?\s+(?:\S+)=|[>"']))+.)?/g);
  // let attachments = await client
  //   .api(`/me/messages/${msgId}/attachments`)
  //   .get();

  // console.log(attachments);

  // attachments.value.map(a => {
  //   msgContent = msgContent.replace(`cid:${a.contentId}`, `data:image/jpeg;base64,${a.contentBytes}`);
  // })

  // for (const a of attachments) {
  //   msgContent = msgContent.replace(`cid:${a.contentId}`, `data:image/jpeg;base64,${a.contentBytes}`);
  // }

  // events.value[0].body.content = msgContent;

  return events;
}

// // id: "AQMkAGQ1YzlkM2UwLTA1YjYtNGQ1MC1hYjAxLTM0Mzc4Mzg2MGIxNwBGAAADeknqoFc1Xkuk6w0DaaGBMQcA-PQOKHH7UE2oBrZnPytBVwAAAgEMAAAA-PQOKHH7UE2oBrZnPytBVwAB6zcZbAAAAA=="
// export async function getAttachments(accessToken, msgId) {
//   const client = getAuthenticatedClient(accessToken);

//   msgId = "AQMkAGQ1YzlkM2UwLTA1YjYtNGQ1MC1hYjAxLTM0Mzc4Mzg2MGIxNwBGAAADeknqoFc1Xkuk6w0DaaGBMQcA-PQOKHH7UE2oBrZnPytBVwAAAgEMAAAA-PQOKHH7UE2oBrZnPytBVwAB6zcZbAAAAA==";

//   const events = await client
//     .api(`/me/messages/${msgId}/attachments`)
//     .get();

//   return events;
// }

// Message message = await graphClient.Me.Messages[id].Request(requestOptions).WithUserAccount(ClaimsPrincipal.Current.ToGraphUserAccount()).Expand("attachments").GetAsync();

export async function getUserPicture(accessToken) {
  const client = getAuthenticatedClient(accessToken);

  const events = await client
  // /users/{id | userPrincipalName}/photo
    // .api('/users/Hugo.Ocampo@endava.com/photo')
    .api('/me/contacts/hugo.ocampo@endava.com/photo')
    // .select('subject,bodyPreview,body,sender')
    .get();

  return events;
}