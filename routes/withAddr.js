var express = require('express');
var router = express.Router();
var excel = require('excel4node');
var phantom = require('phantom');
const jsdom = require("jsdom");
const { JSDOM } = jsdom;

var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet('Sheet 1');
var style = workbook.createStyle({
  font: {
    color: 'blue',
    size: 12
  },
  numberFormat: '$#,##0.00; ($#,##0.00); -'
});
worksheet.cell(1, 1).string('번호').style(style);
worksheet.cell(1, 2).string('학교').style(style);
worksheet.cell(1, 3).string('주소').style(style);

let isParsing = false;
let sheetColCount = 2;
let sheetColNumber = 1;
router.get('/', async function (req, res, next) {
  if (!isParsing) {
    isParsing = true;
    console.log('crawling start!');
    await getSchoolList();
    // await getMiddleSchoolList();
    // await getHighSchoolList();

    const exportExcelName = 'school_list.xlsx';
    workbook.write(exportExcelName);
    console.log('crawling end!');
    // const exportExcelPath = process.cwd();
    // res.download(`${exportExcelPath}/${exportExcelName}`, exportExcelName);
  } else {
    console.log('reject!');
    res.send('reject!');
  }
});

async function getDocumentByUrl(url) {
  const instance = await phantom.create();
  const page = await instance.createPage();
  await page.open(url);
  const content = await page.property('content');
  await instance.exit();
  const { document } = new JSDOM(content).window;
  return document;
}

const getHtmlParam = pageNum => pageNum === 1 ? '.html' : `,1,list01,${pageNum}.html`;
const getElementText = el => el['innerText' in el ? 'innerText' : 'textContent'];

async function getDataAndSetInExcel(baseURL, nameTdIndex, addrTdIndex, locationText, isAddrText) {

  console.log(locationText, 'parsing start', sheetColNumber);
  let document = await getDocumentByUrl(baseURL + getHtmlParam(1));
  // get page total
  const pageWrapEl = document.getElementsByClassName('paging')[0];
  const pageButtonEl = pageWrapEl.getElementsByTagName('a');
  const pageTotal = pageButtonEl.length - 3; // << < > >> 버튼을 제외하고, 1페이지 +

  console.log('total page:', pageTotal);
  for (let pageNum = 1; pageNum <= pageTotal; pageNum++) {
    console.log('pageNum:', pageNum, 'parsing...');
    document = await getDocumentByUrl(baseURL + getHtmlParam(pageNum));
    const trElements = document.getElementsByTagName('tr');
    for (let i = 1; i < trElements.length; i++) {
      const currentTrEl = trElements[i].children;

      const nameTdEl = currentTrEl[nameTdIndex];
      const nameVal = getElementText(nameTdEl);
      const schoolName = String(nameVal).replace('서울', '').replace('등학교', '').replace('학교', '');
      if (schoolName.includes('유치원')) {
        continue;
      }
      const addrTdEl = currentTrEl[addrTdIndex];
      const schoolAddr = isAddrText ? getElementText(addrTdEl) : addrTdEl.children[0].getAttribute('href');

      console.log(sheetColNumber, schoolName, schoolAddr);
      worksheet.cell(sheetColCount, 1).number(sheetColNumber);
      worksheet.cell(sheetColCount, 2).string(schoolName);
      worksheet.cell(sheetColCount, 3).string(schoolAddr);
      sheetColCount++;
      sheetColNumber++;
    }
  }
  console.log(locationText, 'parsing end');
}

async function getDataAndSetInExcelForNoPage(baseURL, nameTdIndex, addrTdIndex, locationText) {
  console.log(locationText, 'parsing start', sheetColNumber);
  let document = await getDocumentByUrl(baseURL);
  const trElements = document.getElementsByTagName('tr');
  for (let i = 2; i < trElements.length; i++) {
    const currentTrEl = trElements[i].children;

    let nameTdEl = currentTrEl[nameTdIndex];
    let schoolName = getElementText(nameTdEl).trim();
    let addrTdEl;
    if (schoolName === '공립' || schoolName === '사립') {
      nameTdEl = currentTrEl[nameTdIndex + 1];
      schoolName = getElementText(nameTdEl).trim();
      addrTdEl = currentTrEl[addrTdIndex + 1];
    } else {
      addrTdEl = currentTrEl[addrTdIndex];
    }
    const schoolAddr = getElementText(addrTdEl);

    console.log(sheetColNumber, schoolName, schoolAddr);
    worksheet.cell(sheetColCount, 1).number(sheetColNumber);
    worksheet.cell(sheetColCount, 2).string(schoolName);
    worksheet.cell(sheetColCount, 3).string(schoolAddr);
    sheetColCount++;
    sheetColNumber++;
  }
  console.log(locationText, 'parsing end');
}

async function getSchoolList() {
  console.log('북부교육청 start');
  await getDataAndSetInExcel('http://bbedu.sen.go.kr/CMS/adminfo/adminfo05/adminfo0502/index', 2, 4, '북부교육청 초', 1);
  await getDataAndSetInExcel('http://bbedu.sen.go.kr/CMS/adminfo/adminfo05/adminfo0503/index', 2, 4, '북부교육청 중', 1);
  await getDataAndSetInExcel('http://bbedu.sen.go.kr/CMS/adminfo/adminfo05/adminfo0504/index', 2, 4, '북부교육청 고', 1);
  console.log('북부교육청 end');
  sheetColCount++;
  console.log('동부교육청 start');
  await getDataAndSetInExcel('http://dbedu.sen.go.kr/CMS/introduction/introduction06/introduction0601/introduction060103/index', 1, 6, '동부교육청 초');
  await getDataAndSetInExcel('http://dbedu.sen.go.kr/CMS/introduction/introduction06/introduction0601/introduction060104/index', 1, 6, '동부교육청 중');
  await getDataAndSetInExcel('http://dbedu.sen.go.kr/CMS/introduction/introduction06/introduction0601/introduction060105/index', 1, 6, '동부교육청 고');
  console.log('동부교육청 end');
  sheetColCount++;
  console.log('성동광진교육청 start');
  await getDataAndSetInExcel('http://sdgjedu.sen.go.kr/CMS/infoedu/infoedu04/infoedu0401/infoedu040102/index', 3, 5, '성동교육청 초');
  await getDataAndSetInExcel('http://sdgjedu.sen.go.kr/CMS/infoedu/infoedu04/infoedu0401/infoedu040103/index', 3, 5, '성동교육청 중');
  await getDataAndSetInExcel('http://sdgjedu.sen.go.kr/CMS/infoedu/infoedu04/infoedu0401/infoedu040104/index', 3, 5, '성동교육청 고');
  console.log('성동광진교육청 end');
  sheetColCount++;
  console.log('성북강북교육청 start');
  await getDataAndSetInExcel('http://sbgbedu.sen.go.kr/CMS/introduction/introduction04/introduction0401/index', 1, 2, '성북강북교육청', 1);
  console.log('성북강북교육청 end');
  sheetColCount++;
  console.log('의정부교육청 start');
  await getDataAndSetInExcelForNoPage('http://www.goeujb.kr/nuri/etc/sub_page.php?pidx=1341907967277', 0, 1, '의정부초등학교');
  await getDataAndSetInExcelForNoPage('http://www.goeujb.kr/nuri/etc/sub_page.php?pidx=1341907989278', 0, 1, '의정부중학교');
  await getDataAndSetInExcelForNoPage('http://www.goeujb.kr/nuri/etc/sub_page.php?pidx=1341908006279', 0, 1, '의정부고등학교');
  console.log('의정부교육청 end');
  sheetColCount++;
}

module.exports = router;
