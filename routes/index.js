var express = require("express");
var router = express.Router();
var excel = require("excel4node");
var phantom = require("phantom");
const jsdom = require("jsdom");
const { JSDOM } = jsdom;

var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet("Sheet 1");
var style = workbook.createStyle({
  font: {
    color: "blue",
    size: 12,
  },
  numberFormat: "$#,##0.00; ($#,##0.00); -",
});
worksheet.cell(1, 1).string("번호").style(style);
worksheet.cell(1, 2).string("학교").style(style);
worksheet.cell(1, 3).string("주소").style(style);

let isParsing = false;
let sheetColCount = 2;
// let sheetColNumber = 1;
router.get("/", async function (req, res, next) {
  if (!isParsing) {
    isParsing = true;
    console.log("crawling start!");
    await getElementarySchoolList();
    await getMiddleSchoolList();
    await getHighSchoolList();

    const exportExcelName = "학교리스트.xlsx";
    workbook.write(exportExcelName);
    console.log("crawling end!");
    // const exportExcelPath = process.cwd();
    // res.download(`${exportExcelPath}/${exportExcelName}`, exportExcelName);
  } else {
    console.log("reject!");
    res.send("reject!");
  }
});

async function getDocumentByUrl(url) {
  const instance = await phantom.create();
  const page = await instance.createPage();
  await page.open(url);
  const content = await page.property("content");
  await instance.exit();
  const { document } = new JSDOM(content).window;
  return document;
}

function getSchoolName({ trElements, i, nameTdIndex, isNameInATag }) {
  const tdEl = trElements[i].getElementsByTagName("td")[nameTdIndex];
  const v = tdEl.textContent;
  const schoolName = String(v)
    .replace("서울", "")
    .replace("등학교", "")
    .replace("학교", "");
  return schoolName;
}

function getSchoolAddr({ trElements, i, addrTdIndex, isAddrInATag }) {
  const tdEl = trElements[i].getElementsByTagName("td")[addrTdIndex];
  return tdEl.textContent;
}

async function getDataAndSetInExcel(
  getUrlByPageNum,
  { locationText, nameTdIndex, isNameInATag, addrTdIndex, isAddrInATag }
) {
  console.log(locationText, "parsing start", sheetColCount - 1);
  let document = await getDocumentByUrl(getUrlByPageNum(1));
  // get page total
  const pageWrapEl = document.getElementsByClassName("paging")[0];
  const pageButtonEl = pageWrapEl.getElementsByTagName("a");
  const pageTotal = pageButtonEl.length - 3; // << < > >> 버튼을 제외하고, 1페이지 +

  console.log("total page:", pageTotal);
  for (let pageNum = 1; pageNum <= pageTotal; pageNum++) {
    console.log("pageNum:", pageNum, "parsing...");
    document = await getDocumentByUrl(getUrlByPageNum(pageNum));
    const trElements = document.getElementsByTagName("tr");
    for (let i = 1; i < trElements.length; i++) {
      const schoolName = getSchoolName({
        trElements,
        i,
        nameTdIndex,
        isNameInATag,
      });
      const schoolAddr = getSchoolAddr({
        trElements,
        i,
        addrTdIndex,
        isAddrInATag,
      });
      console.log(sheetColCount, schoolName, schoolAddr);
      worksheet.cell(sheetColCount, 1).number(sheetColCount - 1);
      worksheet.cell(sheetColCount, 2).string(schoolName);
      worksheet.cell(sheetColCount, 3).string(schoolAddr);
      sheetColCount++;
    }
  }
  console.log(locationText, "parsing end");
}

const OPTION = {
  gsycedu: {
    locationText: "강서양천",
    nameTdIndex: 2,
    isNameInATag: false,
    addrTdIndex: 3,
    isAddrInATag: false,
  },
  dgedu: {
    locationText: "동작관악",
    nameTdIndex: 1,
    isNameInATag: true,
    addrTdIndex: 5,
    isAddrInATag: false,
  },
  sbedu: {
    locationText: "서부",
    nameTdIndex: 1,
    isNameInATag: false,
    addrTdIndex: 2,
    isAddrInATag: true,
  },
  nbedu: {
    locationText: "남부",
    nameTdIndex: 2,
    isNameInATag: false,
    addrTdIndex: 4,
    isAddrInATag: false,
  },
  jbedu: {
    locationText: "중부",
    nameTdIndex: 2,
    isNameInATag: false,
    addrTdIndex: 3,
    isAddrInATag: false,
  },
};

async function getElementarySchoolList() {
  console.log("초등학교 start");
  await getDataAndSetInExcel(
    (pageNum) =>
      pageNum === 1
        ? "http://gsycedu.sen.go.kr/CMS/education/education01/education0102/index.html"
        : `http://gsycedu.sen.go.kr/CMS/education/education01/education0102/index,1,list01,${pageNum}.html`,
    OPTION.gsycedu
  );
  await getDataAndSetInExcel(
    (pageNum) =>
      pageNum === 1
        ? "http://dgedu.sen.go.kr/CMS/introduction/introduction06/introduction0601/introduction060102/index.html"
        : `http://dgedu.sen.go.kr/CMS/introduction/introduction06/introduction0601/introduction060102/index,1,list01,${pageNum}.html`,
    OPTION.dgedu
  );
  await getDataAndSetInExcel(
    (pageNum) =>
      pageNum === 1
        ? "http://sbedu.sen.go.kr/CMS/openedu/openedu01/openedu0103/index.html"
        : `http://sbedu.sen.go.kr/CMS/openedu/openedu01/openedu0103/index,1,list01,${pageNum}.html`,
    OPTION.sbedu
  );
  await getDataAndSetInExcel(
    (pageNum) =>
      pageNum === 1
        ? "http://nbedu.sen.go.kr/CMS/introduction/introduction05/introduction0501/introduction050103/index.html"
        : `http://nbedu.sen.go.kr/CMS/introduction/introduction05/introduction0501/introduction050103/index,1,list01,${pageNum}.html`,
    OPTION.nbedu
  );
  await getDataAndSetInExcel(
    (pageNum) =>
      pageNum === 1
        ? "http://jbedu.sen.go.kr/CMS/introduction/introduction05/introduction0502/index.html"
        : `http://jbedu.sen.go.kr/CMS/introduction/introduction05/introduction0502/index,1,list01,${pageNum}.html`,
    OPTION.jbedu
  );
  console.log("초등학교 end");
}

async function getMiddleSchoolList() {
  console.log("중학교 start");
  await getDataAndSetInExcel(
    (pageNum) =>
      pageNum === 1
        ? "http://gsycedu.sen.go.kr/CMS/education/education01/education0103/index.html"
        : `http://gsycedu.sen.go.kr/CMS/education/education01/education0103/index,1,list01,${pageNum}.html`,
    OPTION.gsycedu
  );
  await getDataAndSetInExcel(
    (pageNum) =>
      pageNum === 1
        ? "http://dgedu.sen.go.kr/CMS/introduction/introduction06/introduction0601/introduction060103/index.html"
        : `http://dgedu.sen.go.kr/CMS/introduction/introduction06/introduction0601/introduction060103/index,1,list01,${pageNum}.html`,
    OPTION.dgedu
  );
  await getDataAndSetInExcel(
    (pageNum) =>
      pageNum === 1
        ? "http://sbedu.sen.go.kr/CMS/openedu/openedu01/openedu0104/index.html"
        : `http://sbedu.sen.go.kr/CMS/openedu/openedu01/openedu0104/index,1,list01,${pageNum}.html`,
    OPTION.sbedu
  );
  await getDataAndSetInExcel(
    (pageNum) =>
      pageNum === 1
        ? "http://nbedu.sen.go.kr/CMS/introduction/introduction05/introduction0501/introduction050104/index.html"
        : `http://nbedu.sen.go.kr/CMS/introduction/introduction05/introduction0501/introduction050104/index,1,list01,${pageNum}.html`,
    OPTION.nbedu
  );
  await getDataAndSetInExcel(
    (pageNum) =>
      pageNum === 1
        ? "http://jbedu.sen.go.kr/CMS/introduction/introduction05/introduction0503/index.html"
        : `http://jbedu.sen.go.kr/CMS/introduction/introduction05/introduction0503/index,1,list01,${pageNum}.html`,
    OPTION.jbedu
  );
  console.log("중학교 end");
}

async function getHighSchoolList() {
  console.log("고등학교 start");
  await getDataAndSetInExcel(
    (pageNum) =>
      pageNum === 1
        ? "http://gsycedu.sen.go.kr/CMS/education/education01/education0104/index.html"
        : `http://gsycedu.sen.go.kr/CMS/education/education01/education0104/index,1,list01,${pageNum}.html`,
    OPTION.gsycedu
  );
  await getDataAndSetInExcel(
    (pageNum) =>
      pageNum === 1
        ? "http://dgedu.sen.go.kr/CMS/introduction/introduction06/introduction0601/introduction060104/index.html"
        : `http://dgedu.sen.go.kr/CMS/introduction/introduction06/introduction0601/introduction060104/index,1,list01,${pageNum}.html`,
    OPTION.dgedu
  );
  await getDataAndSetInExcel(
    (pageNum) =>
      pageNum === 1
        ? "http://sbedu.sen.go.kr/CMS/openedu/openedu01/openedu0105/index.html"
        : `http://sbedu.sen.go.kr/CMS/openedu/openedu01/openedu0105/index,1,list01,${pageNum}.html`,
    OPTION.sbedu
  );
  await getDataAndSetInExcel(
    (pageNum) =>
      pageNum === 1
        ? "http://nbedu.sen.go.kr/CMS/introduction/introduction05/introduction0501/introduction050105/index.html"
        : `http://nbedu.sen.go.kr/CMS/introduction/introduction05/introduction0501/introduction050105/index,1,list01,${pageNum}.html`,
    OPTION.nbedu
  );
  await getDataAndSetInExcel(
    (pageNum) =>
      pageNum === 1
        ? "http://jbedu.sen.go.kr/CMS/introduction/introduction05/introduction0504/index.html"
        : `http://jbedu.sen.go.kr/CMS/introduction/introduction05/introduction0504/index,1,list01,${pageNum}.html`,
    OPTION.jbedu
  );
  console.log("고등학교 end");
}

module.exports = router;
