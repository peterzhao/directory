# How to print the directory
1. Make sure there are Directory-Chinese.csv and Directory-English.csv in the working directory
1. Run `node genereate-html.js` or `NO_ADDRESS=true node genereate-html.js` if you don't want to print address in contact section.
1. open directory.html in Google Chrome
1. Open File/Print in the browser. Make sure:
  - destination: PDF
  - Paper Size: A5
  - Check Footer and Header
  - Margins customized to Minimum but adjust it to allow footer show up.
1. Click Save button to save the file to a PDF file.
1. Open the PDF in Acrobat Reader
1. Open File/Print
1. Make sure:
  - Change Orientation to Landscape
  - Page Size & Handling -> Mutilple -> Pages per sheet: 2 by 1; Page order: Horizontal
1. Click Print button to print
  
