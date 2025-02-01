// Replace these with your actual values
const NOTION_API_KEY = 'your-notion-api-key';
const NOTION_DATABASE_ID = 'your-database-id';
const SHEET_NAME = 'Notion Data';

function updateSheetFromNotion() {
  // Initialize the spreadsheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME) || 
    SpreadsheetApp.getActiveSpreadsheet()
    .insertSheet(SHEET_NAME);
  
  // Clear existing content
  sheet.clear();
  
  // Fetch data from Notion
  const notionData = fetchNotionDatabase();
  
  // Process and write data to sheet
  writeDataToSheet(sheet, notionData);
}

function fetchNotionDatabase() {
  const url = `https://api.notion.com/v1/databases/${NOTION_DATABASE_ID}/query`;
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${NOTION_API_KEY}`,
      'Notion-Version': '2022-06-28',
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  
  return processNotionResponse(data);
}

function processNotionResponse(data) {
  if (!data.results || !data.results.length) {
    throw new Error('No data found in Notion database');
  }
  
  // Extract headers from the first result
  const headers = Object.keys(data.results[0].properties);
  
  // Process each row
  const rows = data.results.map(result => {
    const row = {};
    headers.forEach(header => {
      const property = result.properties[header];
      row[header] = extractPropertyValue(property);
    });
    return row;
  });
  
  return {
    headers: headers,
    rows: rows
  };
}

function extractPropertyValue(property) {
  // Handle different Notion property types
  switch(property.type) {
    case 'title':
      return property.title[0]?.plain_text || '';
    case 'rich_text':
      return property.rich_text[0]?.plain_text || '';
    case 'number':
      return property.number?.toString() || '';
    case 'select':
      return property.select?.name || '';
    case 'multi_select':
      return property.multi_select?.map(item => item.name).join(', ') || '';
    case 'date':
      return property.date?.start || '';
    case 'checkbox':
      return property.checkbox?.toString() || '';
    case 'url':
      return property.url || '';
    case 'email':
      return property.email || '';
    case 'phone_number':
      return property.phone_number || '';
    case 'relation':
      return property.relation?.map(item => item.id).join(', ') || '';
    case 'formula':
      const formulaType = property.formula.type;
      switch(formulaType) {
        case 'string':
          return property.formula.string || '';
        case 'number':
          return property.formula.number?.toString() || '';
        case 'boolean':
          return property.formula.boolean?.toString() || '';
        case 'date':
          return property.formula.date?.start || '';
        default:
          return '';
      }
    default:
      return '';
  }
}

function writeDataToSheet(sheet, data) {
  // Write headers
  sheet.getRange(1, 1, 1, data.headers.length)
    .setValues([data.headers])
    .setFontWeight('bold');
  
  // Write data rows
  if (data.rows.length > 0) {
    const values = data.rows.map(row => 
      data.headers.map(header => row[header])
    );
    
    sheet.getRange(2, 1, values.length, data.headers.length)
      .setValues(values);
  }
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, data.headers.length);
}

// Optional: Add a menu item to run the update
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Notion Sync')
    .addItem('Update from Notion', 'updateSheetFromNotion')
    .addToUi();
}
