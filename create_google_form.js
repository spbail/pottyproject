/*
  Takes tabular raw data with a header and returns a list of records
*/
function sheetAsRecs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Raw Data');
  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var output = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < header.length; j++) {
      row[header[j]] = data[i][j];
    }
    output.push(row);
  }
  return output;
}

/*
  Takes a list of records and groups them in nested hashmaps,
  based on the list of column names.
*/
function groupBy(recs, cols) {
  if (cols.length == 0) {
    return recs;
  }

  var output = {};
  var col = cols.shift();

  var groups = {};
  for (const rec of recs) {
    var val = rec[col];
    if (!(val in groups)) {
      groups[val] = [];
    }
    groups[val].push(rec);
  }

  for (var k in groups) {
    output[k] = groupBy(groups[k], [...cols]);
  }

  return output;
}

/*
  Uses a sheet to persist metadata about state of the current form,
  since we can't generate the entire thing at once without timing out.
*/
function metadata(loc) {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('form metadata')
    .getRange(loc);
}

/*
  Return the current form, or creates a new one if none.
*/
function getForm() {
  var formIdCell = metadata('B1');

  var formId = formIdCell.getValue();
  if (formId != '') {
    return FormApp.openById(formId);
  } else {
    var form = FormApp.create('Potty Project Data Entry');
    formIdCell.setValue(form.getId());
    return form;
  }
}
/*
  getter/setters for the last added page to the form
*/
function getLastDoneBorough() {
  return metadata('B2').getValue();
}
function setLastDoneBorough(val) {
  return metadata('B2').setValue(val);
}
function getLastDonePark() {
  return metadata('B3').getValue();
}
function setLastDonePark(val) {
  return metadata('B3').setValue(val);
}

/*
  Apps scripts time out after about 5/6 min,
  we can be conservative and exit gracefully after 4.
*/
function isTimeUp() {
  var now = new Date();
  return now.getTime() - executionStart.getTime() > 240000; // 4 minutes
}

/*
  Create the section for actual data entry for specific park/potty.
*/
function createParkSection(form, parkName, borough, recs) {
  var section = getParkSection(form, parkName, borough);

  // Need to de-dup the names, apparently we have some duplicates,
  // which breaks the multiple choice item.
  form.addMultipleChoiceItem()
    .setTitle('Potty Name')
    .setChoiceValues([... new Set(recs.map(r => r['Name']))]);

  form.addMultipleChoiceItem()
    .setTitle('Type')
    .setChoiceValues(['Permanent', 'Portapotty', 'Trailer']);

  form.addScaleItem()
    .setTitle('Number of Stalls')
    .setBounds(1, 6)
    .setLabels('', 'or more');

  form.addTimeItem()
    .setTitle('Opening time');

  form.addTimeItem()
    .setTitle('Closing Time');

  form.addParagraphTextItem()
    .setTitle('Any other info');

  return section;
}

function getBoroughChoice(form) {
  var lists = form.getItems(FormApp.ItemType.LIST);

  var choice = lists.find(e => e.getTitle() === 'Borough');
  if (choice) {
    return choice.asListItem();
  }
  return form.addListItem()
    .setTitle('Borough');
}
function getParkChoice(form, borough) {
  var lists = form.getItems(FormApp.ItemType.LIST);

  var choice = lists.find(e => e.getTitle() === 'Park Name' && e.getHelpText() === borough);
  if (choice) {
    return choice.asListItem();
  }
  return form.addListItem()
    .setTitle('Park Name')
    .setHelpText(borough);
}

function getBoroughSection(form, borough) {
  var sections = form.getItems(FormApp.ItemType.PAGE_BREAK);
  var section = sections.find(e => e.getTitle() === borough && e.getHelpText() === '');
  if (section) {
    return section.asPageBreakItem();
  }
  return form.addPageBreakItem()
    .setTitle(borough)
    .setGoToPage(FormApp.PageNavigationType.SUBMIT);
}

function getParkSection(form, park, borough) {
  var sections = form.getItems(FormApp.ItemType.PAGE_BREAK);
  var section = sections.find(e => e.getTitle() === park && e.getHelpText() === borough);
  if (section) {
    return section.asPageBreakItem();
  }
  return form.addPageBreakItem()
    .setTitle(borough)
    .setGoToPage(FormApp.PageNavigationType.SUBMIT);
}

/*
  Create the entry point and each follow-up park section for a borough.
*/
function createBoroughSection(form, borough, parksAndRecs) {
  var section = getBoroughSection(form, borough);

  var parkChoice = getParkChoice(form, borough);

  var parks = Object.keys(parksAndRecs);
  parks.sort();

  for (const p of parks) {
    if (isTimeUp()) {
      throw Error('Time is up'); // exit semi-gracefully
    }

    if (p <= getLastDonePark()) {
      continue;
    }

    Logger.log(p);

    createParkSection(form, p, borough, parksAndRecs[p]);

    setLastDonePark(p);
  }

  parkChoice.setChoices(
    parks.map(p => parkChoice.createChoice(p, getParkSection(form, p, borough)))
  );

  setLastDonePark('');
  return section;
}

/*
  Main function to create a form based on the raw potty data sheet
*/
function createPottyForm(groupedData) {
  var form = getForm();
  var boroughs = Object.keys(groupedData);
  boroughs.sort();

  var boroughChoice = getBoroughChoice(form);

  for (const b of boroughs) {
    if (isTimeUp()) {
      throw Error('Time is up'); // exit semi-gracefully
    }

    if (b <= getLastDoneBorough()) {
      continue;
    }

    Logger.log(b);

    createBoroughSection(form, b, groupedData[b]);

    setLastDoneBorough(b);
  }

  boroughChoice.setChoices(
    boroughs.map(b => boroughChoice.createChoice(b, getBoroughSection(form, b)))
  );
}

var executionStart = new Date();

createPottyForm(groupBy(sheetAsRecs(), ["Borough", "Park Name"]));
