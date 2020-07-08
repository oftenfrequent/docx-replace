const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const fs = require('fs');
const path = require('path');
const donerlist = fs.readFileSync(path.resolve(__dirname, './files/donerlist.csv'), 'utf8');
const content = fs.readFileSync(path.resolve(__dirname, 'files/template.docx'), 'binary');

console.log('starting')

if (!fs.existsSync(path.resolve(__dirname, './output'))){
  console.log('created output folder')
  fs.mkdirSync(path.resolve(__dirname, './output'));
}

const createDocPerDoner = async (doc) => {
  // split by rows
  await donerlist.split('\n').reduce(async (promiseChain, doner) => {
    // wait for each doc to be created before moving on
    await promiseChain.then(async () => {
      // split only on commas without spaces after
      const info = doner.split(/,(?!\s)/);
      if (info[1].indexOf('$') > -1) {
        const data = {
          doner: info[0].replace('"', '').replace('"', ''),
          amount: info[1]
        }
        doc.setData(data);
        try {
          doc.render()
          const buf = doc.getZip().generate({type: 'nodebuffer'});
          fs.writeFileSync(path.resolve(__dirname, `./output/${data.doner.replace('/', '')}-thank-you-2020.docx`), buf);
          console.log(`Finished: ${data.doner}`)
        } catch (err) {
          throw err;
        }
      }
    });
  }, Promise.resolve())

}

const zip = new PizZip(content);
try {
  const doc = new Docxtemplater(zip);
  createDocPerDoner(doc)
} catch(error) {
  console.error(error);
}
