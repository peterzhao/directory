const fs = require('fs');
const os = require('os');
const Excel = require('exceljs');

const sheetNames = ['Chinese', 'English'];

const createHtml = (data) => {
  const head = `
<head>
    <meta charset="UTF-8">
    <style>
    .col-container {
      display: table; 
      width: 100%;
     }

     .col {
       display: table-cell;
       padding: 3px;
     }
     
     li {
       page-break-inside: avoid;
       border-style: solid;
       border-width: 1px;  
       margin-bottom: 8px;
     }
     
     ul {
       list-style-type: none;
     }
     
     div {
       font-size: x-small;
     }
     .name {
       border-right-style: solid;
       border-right-width: 1px;  
       width: 58%;  
     }
     .family-name {
       width: 29%;
       font-weight: bold;
     }
     .members {
       width: 70%;
     }
     .first-name {
       width: 55%;
     }
     .chinese-name {
       width: 44%;
     }
     .contact {
       width: 40%;
       padding-left: 10px;
     }
     .phone-number {
       font-weight: bold;
     }
     .section-header {
       font-weight: bold;
       font-size: medium;
       text-align: center;
       page-break-before: always;
       margin: 10px;
     }
    </style>
  </head>
`;

  const html = `<html>
  ${head}
  <body>
    <ul>
      ${data}
    </ul>
  </body>
</html>
`;

  fs.writeFileSync('directory.html', html);
};

const notEmpty = (text) => {
  return text && text.trim && text.trim().length > 0;
}

function getFamilies(worksheet) {
  const families = [];
  let family;
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const familyName = row.values[1] || '';
    const firstName = row.values[2] || '';
    const chineseName = row.values[3] || '';
    const address = row.values[4] || '';
    const city = row.values[5] || '';
    const zipCode = row.values[6] || '';
    const phoneNumber = row.values[7] || '';
    const email = row.values[8] || '';
    if (notEmpty(familyName)) {
      family = {familyName, address, cityZipCode: `${city} ${zipCode}`, phoneNumber, email};
      family.members = [];
      families.push(family);
    }
    if (family && (notEmpty(familyName) || notEmpty(firstName) || notEmpty(chineseName))) {
      family.members.push({firstName, chineseName});
    }
  });
  return families;
}

function createDivs(data, className) {
  if (!data) return '';
  if (['address', 'cityZipCode'].includes(className) && process.env.NO_ADDRESS === 'true') return '';
  return data.split(os.EOL).map((value) => `<div class="${className}">${value}</div>`).join(os.EOL);
}

function createDirectoryHtml(families) {
  return families.map((family) => {
    const members = family.members.map((member) => {
      return `<div class="member col-container">
              <div class="first-name col">${member.firstName}</div>
              <div class="chinese-name col">${member.chineseName}</div>
            </div>`
    }).join(os.EOL);
    return `
         <li class="col-container">
        <div class="name col">
          <div class="col-container">
            <div class="family-name col">${family.familyName}</div>
            <div class="members col">
              ${members}
            </div>
          </div>
        </div>
        <div class="contact col">
          ${createDivs(family.address, 'address')}
          ${createDivs(family.cityZipCode, 'cityZipCode')}
          ${createDivs(family.phoneNumber, 'phone-number')}
          ${createDivs(family.email, 'email')}
        </div>
      </li>
        `;
  });
}

Promise.all(sheetNames.map((sheetName) => {
  var workbook = new Excel.Workbook();
  return workbook.csv.readFile(`Directory-${sheetName}.csv`)
    .then((worksheet) => {
      const families = getFamilies(worksheet);
      return createDirectoryHtml(families);
    });
}))
  .then((elements) => {
    createHtml(`
      ${elements[0].join(os.EOL)}
      <div class="section-header">English Ministry</div>
      ${elements[1].join(os.EOL)}
      `);
  });
