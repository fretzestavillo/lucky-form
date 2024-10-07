"use strict";

// Get new data from Sheet
function doGet() {
  const ss = SpreadsheetApp.openById(
    "191-7sXPiff4ljNuJ-KgLupkezUDicdWMDKKqF0E4O5E"
  );
  const ws = ss.getSheetByName("Party Ulam");
  if (!ws) return;
  const data = ws.getDataRange().getValues();
  const values = data[data.length - 1];
  const headers = ["timeStamp", "name", "email", "street", "id"];
  const result = values.map((value, i) => ({
    [headers[i]]: value,
  }));
  const obj = Object.assign({}, ...result);
  createDocument(obj);
}

// Create google document and send to email
function createDocument(obj) {
  const doc = DocumentApp.create("secret file 69");
  const body = doc.getBody();
  const url = `https://fretzestavillo.github.io/sura-generator-gh-page/?id=${obj.id}`;
  console.log(url);
  console.log(obj.email, "creating document");

  body.appendParagraph(`

Dear ${obj.name},

Thank you for signing up! We are excited to have you participate in this opportunity to win amazing prizes.

To proceed and check if you are one of the lucky winners, please click the link below:
`);
  const linkText = body.appendParagraph(
    "Click here to check if you're a lucky winner!"
  );
  linkText.setLinkUrl(url);
  body.appendParagraph(`
We wish you the best of luck and hope you find your name among the winners!

Thank you once again for your participation, and we look forward to your continued engagement.

Best regards,
George
CEO
  `);
  doc.saveAndClose();
  sendEmail(doc, obj);
}
function sendEmail(doc, obj) {
  GmailApp.sendEmail(
    `${obj.email}`,
    "Thanks for signing up!!!",
    "Please see the attached file.",
    {
      attachments: doc,
      name: "Google Account Team",
    }
  );
}

// post from sura generator
function doPost(e) {
  const body = e.postData.contents;
  const bodyJSON = JSON.parse(body);
  const suraPrize = bodyJSON;

  if (suraPrize.id === "") return;

  getAlldata(suraPrize);
  console.log(suraPrize.id, suraPrize.sura, "sura generator sent");
  return ContentService.createTextOutput(
    JSON.stringify({ result: bodyJSON })
  ).setMimeType(ContentService.MimeType.JSON);
}

// get all info to check if id is iclude
function getAlldata(suraPrize) {
  const idSura = suraPrize.id;
  const prize = suraPrize.sura;
  const ss = SpreadsheetApp.openById(
    "191-7sXPiff4ljNuJ-KgLupkezUDicdWMDKKqF0E4O5E"
  );
  const ws = ss.getSheetByName("Party Ulam");
  if (!ws) return;
  const getValues = ws.getDataRange().getValues();
  const values = getValues.slice(1);
  const headers = ["timeStamp", "name", "email", "street", "id"];
  const transform = (value) =>
    value.reduce((acc, cur, i) => {
      acc[headers[i]] = cur;
      return acc;
    }, {});

  const data = values.map(transform);
  const obj = data.filter((obj) => {
    return obj.id === idSura;
  });

  const filteredObj = obj[0];
  if (!filteredObj) return;

  console.log(filteredObj.email, "get amail from sheet");
  createMap(filteredObj, prize);
}

// create google map
function createMap(filteredObj, prize) {
  console.log(filteredObj.email, "creating map and pass email");
  const streetAddsource = `${filteredObj.street}`;
  const streetAddTarget = "Sta. Teresita Parish (Archdiocese of Manila)";
  console.log(filteredObj.email);

  const directions = Maps.newDirectionFinder()
    .setOrigin(streetAddsource)
    .setDestination(streetAddTarget)
    .setMode(Maps.DirectionFinder.Mode.DRIVING)
    .getDirections();
  var route = directions.routes[0];

  var markerSize = Maps.StaticMap.MarkerSize.SMALL;
  var markerColor = Maps.StaticMap.Color.RED;
  var markerLetterCode = "A".charCodeAt();
  var map = Maps.newStaticMap();
  for (var i = 0; i < route.legs.length; i++) {
    var leg = route.legs[i];
    if (i == 0) {
      map.setMarkerStyle(
        markerSize,
        markerColor,
        String.fromCharCode(markerLetterCode)
      );
      map.addMarker(leg.start_location.lat, leg.start_location.lng);
      markerLetterCode++;
    }
    map.setMarkerStyle(
      markerSize,
      markerColor,
      String.fromCharCode(markerLetterCode)
    );
    map.addMarker(leg.end_location.lat, leg.end_location.lng);
    markerLetterCode++;
  }

  map.addPath(route.overview_polyline.points);

  // Send the map in an email.
  MailApp.sendEmail(
    `${filteredObj.email}`,
    `Congratulations! You’ve Won ${prize} – Claim Your Prize Today!`,
    "Please open: " + map.getMapUrl() + "&key=YOUR_API_KEY",
    {
      name: "Google Account Team",
      htmlBody: `
Dear ${filteredObj.name},
<br /><br />
Congratulations! You Won ${prize}!
<br /><br />
We are thrilled to inform you that you have been selected as one of the lucky winners in our raffle! Your creativity and effort have truly shone through, and we are excited to reward you for your achievement.
<br /><br />
To claim your prize, please follow the direction below and look for Sister Cassy at the designated location. Kindly make sure to bring a valid ID to verify your identity.
<br /><br />
Once again, congratulations! We look forward to seeing you soon and hope this prize brings you as much joy as your participation brought to us.
<br /><br />
Best regards,
<br />
George Malone
<br />
CEO
<br /><br />
location below:
<br/><img src="cid:mapImage">`,
      inlineImages: {
        mapImage: Utilities.newBlob(map.getMapImage(), "image/png"),
      },
    }
  );

  deleteRecords(filteredObj);
}

function deleteRecords(filteredObj) {
  const currentUser = filteredObj.id;
  console.log("user deleted");

  const SS = SpreadsheetApp.openById(
    "191-7sXPiff4ljNuJ-KgLupkezUDicdWMDKKqF0E4O5E"
  );
  const SHEET = SS.getSheetByName("Party Ulam");
  const RANGE = SHEET.getDataRange();
  const DELETE_VAL = currentUser;
  const COL_TO_SEARCH = 4; // The column to search for the DELETE_VAL (Zero is first)

  var rangeVals = RANGE.getValues();
  for (let i = rangeVals.length - 1; i >= 0; i--) {
    if (rangeVals[i][COL_TO_SEARCH] === DELETE_VAL) {
      SHEET.deleteRow(i + 1);
    }
  }

  deleteDocFile();
}

function deleteDocFile() {
  var fileName = "secret file 69";

  var files = DriveApp.getFilesByName(fileName);

  if (files.hasNext()) {
    var file = files.next();
    file.setTrashed(true);
    console.log("file has been deleted");
  } else {
    console.log("no files found");
  }
}
