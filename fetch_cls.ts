import * as https from 'https';
import * as fs from 'fs';
import * as XLSX from 'xlsx';

function downloadFile(url: string, dest: string): Promise<void> {
  return new Promise((resolve, reject) => {
    https.get(url, (response) => {
      if (response.statusCode === 307 || response.statusCode === 302 || response.statusCode === 301) {
        downloadFile(response.headers.location as string, dest).then(resolve).catch(reject);
      } else {
        const file = fs.createWriteStream(dest);
        response.pipe(file);
        file.on('finish', () => {
          file.close();
          resolve();
        });
      }
    }).on('error', (err) => {
      reject(err);
    });
  });
}

downloadFile('https://docs.google.com/spreadsheets/d/1jmwwVX35habW5LKEFgQpUj1atSCz-LEF/export?format=xlsx', 'cls.xlsx')
  .then(() => {
    const fileData = fs.readFileSync('cls.xlsx');
    const workbook = XLSX.read(fileData, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, {header: 1});
    console.log(JSON.stringify(data.slice(0, 30), null, 2));
  })
  .catch(err => console.error(err));
