const Excel = require('exceljs');
const os = require('os');

const cityZipCodeExpression = /^([a-zA-Z\s]+)[,，]\s*([A-Z][0-9][A-Z]\s*[0-9][A-Z][0-9])$/;
const sheetNames = ['Chinese', 'English'];


const notEmpty = (text) => {
  return text && text.trim && text.trim().length > 0;
};

const getText = (data) => {
  if (!data) return data;
  if (typeof(data) === 'object') {
    if (data.text && data.text.trim) {
      return data.text.trim();
    } else if (data.richText && Array.isArray(data.richText)) {
      return data.richText.map(rt => rt.text).join(' ');
    }
  }
  if(typeof(data) === 'string') return data.trim();
  return data;
}

const isEmail = (data) => {
  const expression = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return expression.test(data);
};

const isPhoneNumber = (data) => {
  const expression = /^\([0-9]{3}\)\s*[0-9]{3}-[0-9]{4}$/;
  return expression.test(data);
};

const isCityZipCode = (data) => {
  return cityZipCodeExpression.test(data);
};

const getCity = (data) => {
  const match = cityZipCodeExpression.exec(data);
  return match[1];
};

const getZipCode = (data) => {
  const match = cityZipCodeExpression.exec(data);
  return match[2];
};

const parseContact = (family, data) => {
  if (!data || !family) return;
  if (isEmail(data)) {
    if (!family.email) {
      family.email = data;
    } else {
      family.email += `${os.EOL}${data}`;
    }
  } else if (isPhoneNumber(data)) {
    if (!family.phoneNumber) {
      family.phoneNumber = data;
    } else {
      family.phoneNumber += `${os.EOL}${data}`;
    }
  } else if (isCityZipCode(data)) {
    family.city = getCity(data);
    family.zipCode = getZipCode(data);
  } else if (!family.address){
    family.address = data;
  } else {
    family.address += `${os.EOL}${data}`;
  }
};

function readSheet(workbook, sheetName) {
  const worksheet = workbook.getWorksheet(sheetName);
  const families = [];
  let family;
  worksheet.eachRow((row, rowNumber) => {
    const familyName = getText(row.values[1]);
    const firstName = getText(row.values[2]);
    const chineseName = getText(row.values[3]);
    if ((notEmpty(familyName) && familyName !== 'English Ministry')
      || rowNumber === (worksheet.rowCount - 1)) {
      if (family) families.push(family);
      family = {}
      family.members = [];
    }
    if (family && (notEmpty(familyName) || notEmpty(firstName) || notEmpty(chineseName))) {
      family.members.push({familyName, firstName, chineseName});
    }
    parseContact(family, getText(row.values[5]));
  });
  return families;
};

const writeToFile = (families, sheetName) => {
  const header = ['Family Name', 'First Name', 'Chinese Name', 'Address', 'City', 'Zip Code', 'Phone Number', 'Email'];
  const rows = [];
  rows.push(header);
  families.forEach((family) => {
    const member = family.members[0];
    rows.push([member.familyName,
      member.firstName,
      member.chineseName,
      family.address,
      family.city,
      family.zipCode,
      family.phoneNumber,
      family.email
    ]);
    for(let i = 1; i < family.members.length; i += 1) {
      rows.push([undefined, family.members[i].firstName, family.members[i].chineseName]);
    }
  });
  const filename = `Directory-${sheetName}.csv`;
  var workbook = new Excel.Workbook();
  var sheet = workbook.addWorksheet(sheetName);
  sheet.addRows(rows);
  workbook.csv.writeFile(filename).then(function() {
    console.log(`New file has been written to ${filename}`);
  });
};

const workbook = new Excel.Workbook();
workbook.xlsx.readFile('directory2016.xlsx')
  .then(function() {
    sheetNames.forEach((sheetName) => {
      writeToFile(readSheet(workbook, sheetName), sheetName);
    });
  });