const request = require('request');
const cheerio = require('cheerio');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
let wikipath = path.join(__dirname, 'Wiki_of_D');
let url = 'https:';

dirCreator(wikipath);
function dirCreator(filepath) {
  if (!fs.existsSync(filepath)) fs.mkdirSync(filepath);
}

request('https://www.wikipedia.org/', maincb);

function maincb(error, response, html) {
  if (error) {
    console.log(error);
  } else {
    mainsource(html);
  }
}

function mainsource(html) {
  let $ = cheerio.load(html);
  url = url + $($('#js-link-box-en')[0]).attr('href');
  url = url.substring(0, url.length - 1);
  request(url, allportalscb);
}

function allportalscb(err, res, html) {
  if (err) {
    console.log(err);
  } else {
    allportalssource(html.toString());
  }
}
function allportalssource(html) {
  let $ = cheerio.load(html);
  let arlist = $('#mp-portals>li>a');
  request(url + $(arlist[arlist.length - 1]).attr('href'), azindexcb);
}
function azindexcb(err, res, html) {
  if (err) {
    console.log(err);
  } else {
    azindexsource(html);
  }
}

function azindexsource(html) {
  let $ = cheerio.load(html);
  let linksarr = $('.hlist.noprint>ul>li>a');
  for (let i = 0; i < linksarr.length; i++) {
    if ($(linksarr[i]).attr('title') == 'Wikipedia:Contents/A–Z index') {
      linksarr = url + $(linksarr[i]).attr('href');
      break;
    }
  }
  request(linksarr, initialscb);
}

function initialscb(err, res, html) {
  if (err) {
    console.log(err);
  } else {
    initialssource(html);
  }
}

function initialssource(html) {
  let $ = cheerio.load(html);
  let linksar = $('#toc>tbody>tr>td>b>a');
  for (let j = 0; j < linksar.length; j++) {
    if ($(linksar[j]).text() == 'D') {
      linksar = $(linksar[j]).attr('href');
      break;
    }
  }
  request(url + linksar, myinitialcb);
}

function myinitialcb(err, res, html) {
  if (err) {
    console.log(err);
  } else {
    myinitialsource(html);
  }
}

function myinitialsource(html) {
  let $ = cheerio.load(html);
  let linksar = $($('.mw-allpages-chunk>li>a')[0]).attr('href');
  request(url + linksar, finalcb);
}

function finalcb(err, res, html) {
  if (err) {
    console.log(err);
  } else {
    finalsource(html);
  }
}
function finalsource(html) {
  let $ = cheerio.load(html);
  let linksar = $('#mw-content-text>.mw-parser-output>');
  let sheetname = '';
  let existdata = [];
  let newdata = [];
  wikipath = path.join(wikipath, 'wikisheet.xlsx');
  for (let i = 0; i < linksar.length; i++) {
    if ($(linksar[i]).find('span').hasClass('mw-headline'))
      sheetname = $(linksar[i]).text().trim();
    if (
      $(linksar[i]).find('a').attr('title') == 'Semitic languages' ||
      $(linksar[i]).find('a').attr('title') ==
        'International Phonetic Alphabet' ||
      $(linksar[i]).find('a').attr('title') == 'Roman numeral'
    ) {
      newdata = $(linksar[i]).text().trim();
      newdata = { Data: newdata };
      existdata = excelReader(wikipath, sheetname);
      existdata.push(newdata);
      if ($(linksar[i]).find('a').attr('title') == 'Semitic languages') {
        newdata = $(linksar[i + 1])
          .text()
          .trim();
        newdata = { Data: newdata };
        existdata.push(newdata);
      }
      excelWriter(wikipath, existdata, sheetname);
    }
  }
}

function excelReader(filepath, sheetname) {
  if (!fs.existsSync(filepath)) return [];
  let workbook = xlsx.readFile(filepath);
  let exceldata = workbook.Sheets[sheetname];
  let data = xlsx.utils.sheet_to_json(exceldata);
  return data;
}

function excelWriter(filepath, data, sheetname) {
  let workbook = '';
  if (!fs.existsSync(filepath)) workbook = xlsx.utils.book_new();
  else workbook = xlsx.readFile(filepath);
  try {
    let worksheet = xlsx.utils.json_to_sheet(data);
    xlsx.utils.book_append_sheet(workbook, worksheet, sheetname);
  } catch (e) {
    console.log(e.message);
  }
  xlsx.writeFile(workbook, filepath);
}
