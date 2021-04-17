const outermostfolderId = ""; // 最外面資料夾ID 如: 2021年天氣學
const sheetId = ""; //google excel id
const dayArray = [20210417,20210418,20210419,20210420,20210421,20210422,20210423] // 你要創的日期 2021年天氣學 > 20210324 ...

const cwbStationID = [ '46692', '46699', '46734', '46750', '46810' ,'46780']; // 可新增/移除測站

const timeMachine = 0; // 輸入數字(若數字>0，則為退回 timeMachine 小時) // 讓程式把時間退回幾小時前，避免有時段的圖沒抓到 (PS 抓圖一樣是一次抓 6hr 的喔!)

// ---------------------------------------------- 下面超醜 不用管
function runThisCodeToCreateFolders(){
  inputTitleToExcel()
  dayArray.forEach(eachDay=>{
    recordID(eachDay);
  })
}

// img 參數  1.類別 2.web( 不同網站網址處理不太一樣，預設 "CWB" ) 3.類別類的細項( 預設 全部 )
function part1(){
    img('雷達');
}

// 因為執行起來有機會超過6min，所以改加個參數選擇你要跑的細項
function part2_1() {
  img('衛星雲圖','CWB',['可見光_台灣','可見光_亞洲']);
}
function part2_2() {
  img('衛星雲圖','CWB',['色調強化_台灣','色調強化_亞洲']);
}
function part2_3() {
  img('衛星雲圖','CWB',['真實色_亞洲','真實色_台灣']);
}

function part3(){
  img('氣溫');
  img('雨量');
}
function part4(){
  img('NCDR風場','NCDR');
}
// part5是後面加的懶得整併了
function part5(){
  img('日累積雨量')
  img_part5(['探空','JMA天氣圖'])
}

// --------------------------------------------------------------------------------------------- folder 專區
function createFolder(folderID, folderName){
  var parentFolder = DriveApp.getFolderById(folderID);
  var subFolders = parentFolder.getFolders();
  var doesntExists = true;
  var newFolder = '';
  
  // Check if folder already exists.
  while(subFolders.hasNext()){
    var folder = subFolders.next();
    
    //If the name exists return the id of the folder
    if(folder.getName() === folderName){
      doesntExists = false;
      newFolder = folder;
      return newFolder.getId();
    };
  };
  //If the name doesn't exists, then create a new folder
  if(doesntExists == true){
    //If the file doesn't exists
    newFolder = parentFolder.createFolder(folderName);
    return newFolder.getId();
  };
};
 
function start(NEW_FOLDER_NAME,parentFolderId){
  var FOLDER_ID = parentFolderId;
  var myFolderID = createFolder(FOLDER_ID, NEW_FOLDER_NAME);
  return myFolderID
};

// ---------------------------------------------------------------------
const imgParams = ['雷達-台灣周邊','雷達-全範圍','衛星雲圖-可見光_台灣','衛星雲圖-可見光_亞洲','衛星雲圖-色調強化_台灣','衛星雲圖-色調強化_亞洲','衛星雲圖-真實色_台灣','衛星雲圖-真實色_亞洲','氣溫','小時累積雨量','日累積雨量','探空','JMA天氣圖','NCDR風場']

function inputTitleToExcel(){
  // // 初始化試算表
  let SpreadSheet = SpreadsheetApp.openById(sheetId);
  let Sheet = SpreadSheet.getSheets()[0]; // 指定第一張試算表
  let LastRow = Sheet.getLastRow(); // 取得最後一列有值的索引值

  // 寫入試算表
  Sheet.getRange(LastRow+1, 1).setValue('日期');
  imgParams.forEach((item,index)=>{
    Sheet.getRange(LastRow+1, index+2).setValue(item);
  })
}
function recordID(date) {
  // // 初始化試算表
  let SpreadSheet = SpreadsheetApp.openById(sheetId);
  let Sheet = SpreadSheet.getSheets()[0]; // 指定第一張試算表
  let LastRow = Sheet.getLastRow(); // 取得最後一列有值的索引值
  
  // 寫入試算表
  Sheet.getRange(LastRow+1, 1).setValue(date);

  //date 為名稱 202-03-24 之類的都可
  const dateFolderId = start(date,outermostfolderId)
  imgParams.forEach((item,index)=>{
    Sheet.getRange(LastRow+1, index+2).setValue(start(item,dateFolderId));
  })
  Logger.log('寫入成功')
}


function getExcelData(e) {
  // 初始化試算表
  let SpreadSheet = SpreadsheetApp.openById(sheetId);
  let Sheet = SpreadSheet.getSheets()[0]; // 指定第一張試算表

  var rowLength = Sheet.getLastRow()-1; //取行長度
  var columnLength = Sheet.getLastColumn(); //取列長度
  var data = Sheet.getRange(2,1,rowLength,columnLength).getValues(); // 取得的資料
  var dataExportId = {};
  
  for(i in data){
    if(data[i][0] != ""){
    dataExportId[data[i][0]] = {
      '台灣周邊':   data[i][1],
      '全範圍':   data[i][2],
      '可見光_台灣':  data[i][3],
      '可見光_亞洲':   data[i][4],
      '色調強化_台灣':   data[i][5],
      '色調強化_亞洲':  data[i][6],
      '真實色_台灣':   data[i][7],
      '真實色_亞洲':   data[i][8],
      '氣溫':  data[i][9],
      '小時累積雨量':   data[i][10],
      '日累積雨量':  data[i][11],
      '探空':  data[i][12],
      'JMA天氣圖':   data[i][13],
      'NCDR風場':  data[i][14]
      }
    }
  }
  return dataExportId;
}

const folderId = function(item){
 const jsonForfolderId = getExcelData();
 return jsonForfolderId[folderDay][item];
}

// --------------------------------------------------------------------------------------------- part5 想截'空氣品質',但失敗
// function part5(presentationId) {
//   var siteUrl = "https://airtw.epa.gov.tw/";
//   var url = "https://pagespeedonline.googleapis.com/pagespeedonline/v5/runPagespeed?screenshot=true&fields=screenshot&url=" + encodeURIComponent(siteUrl);
//   var res = UrlFetchApp.fetch(url).getContentText();
//   var obj = JSON.parse(res);
//   var blob = Utilities.newBlob(Utilities.base64DecodeWebSafe(obj.screenshot.data), "image/png", "sample.png");
//   DriveApp.createFile(blob);
// }
// -----------------------------------------------------------  get 圖片專區
function getImg(url,id){
  try {
  var response  = UrlFetchApp.fetch(url).getBlob();
    var dir =DriveApp.getFolderById(id)
    Logger.log(response)
    dir.createFile(response);
  } catch (error) {
  console.error(error);
  } 
}
// -----------------------------------------------------------------------
let folderDay = Utilities.formatDate(new Date(Date.now()-timeMachine*60*60*1000), "GMT+8", "yyyyMMdd")

const part5Data = {
  '探空': {
    ids: cwbStationID,
    times: [ '00', '12' ],
    url_start: 'https://npd.cwb.gov.tw/NPD/irisme_data/Weather/SkewT/SKW___000_', 
    url_end: '.gif',
  },
  'JMA天氣圖': {
    ids: [ 'aupq78_r', 'aupq35_r' ],
    url_start: 'https://n-kishou.com/ee/image4/lfax/resize/', //https://n-kishou.com/ee/image4/lfax/aupq35_202101010900.png
    times: [ '0900', '2100' ], 
    url_end: '.png'
  }
}

function imgFactory(property) {
  this.imgProperty = {
    雷達: {
      category: {
        台灣周邊: {
          img_url: "https://www.cwb.gov.tw/Data/radar/CV1_TW_3600_"
        },
        全範圍: {
          img_url: "https://www.cwb.gov.tw/Data/radar/CV1_3600_"
        }
      },
      timeFormat: "YYYYMMDD[00-23][0-5]0",
      timeGap: 10,
      imgFormat: ".png"
    },

    衛星雲圖: {
      category: {
        色調強化_台灣: {
          img_url:
            "https://www.cwb.gov.tw/Data/satellite/TWI_IR1_MB_800/TWI_IR1_MB_800-"
        },

        色調強化_亞洲: {
          img_url:
            "https://www.cwb.gov.tw/Data/satellite/LCC_IR1_MB_2750/LCC_IR1_MB_2750-"
        },
        可見光_台灣: {
          img_url:
            "https://www.cwb.gov.tw/Data/satellite/TWI_VIS_Gray_1350/TWI_VIS_Gray_1350-"
        },

        可見光_亞洲: {
          img_url:
            "https://www.cwb.gov.tw/Data/satellite/LCC_VIS_Gray_2750/LCC_VIS_Gray_2750-"
        },
        真實色_台灣: {
          img_url:
            "https://www.cwb.gov.tw/Data/satellite/TWI_VIS_TRGB_1375/TWI_VIS_TRGB_1375-"
        },

        真實色_亞洲: {
          img_url:
            "https://www.cwb.gov.tw/Data/satellite/LCC_VIS_TRGB_2750/LCC_VIS_TRGB_2750-"
        }
      },
      timeFormat: "YYYY-MM-DD-[00-23]-[0-5]0",
      timeGap: 10,
      imgFormat: ".jpg"
    },
    氣溫: {
      category: {
        氣溫: {
          img_url: "https://www.cwb.gov.tw/Data/temperature/"
        }
      },
      timeFormat: "YYYY-MM-DD_[00-23]00",
      timeGap: 60,
      imgFormat: ".GTP8.jpg"
    },
    雨量: {
      category: {
        小時累積雨量: {
          img_url: "https://www.cwb.gov.tw/Data/rainfall/"
        },
      },
      timeFormat: "YYYY-MM-DD_[00-23]00",
      timeGap: 60,
      imgFormat: ".QZT8.jpg"
    },
    日累積雨量: {
      category: {
        日累積雨量: {
          img_url: "https://www.cwb.gov.tw/Data/rainfall/"
        }
      },
      timeFormat: "YYYY-MM-DD_0000",
      timeGap: 0,
      imgFormat: ".QZJ8.jpg"
    },
    NCDR風場: {
      category: {
        NCDR風場: {
          img_url: "https://watch.ncdr.nat.gov.tw/00_Wxmap/5A7_CWB_WINDMAP/"
        }
      },
      timeFormat: "YYYYMMDD[00-23][0-5]0",
      timeGap: 60,
      imgFormat: ".png"
    }
  };
  this.propertyWantToDraw = property; //?????
  this.array;
}

imgFactory.prototype.timeInterval = function (interval_var){
  let today = new Date(Date.now()-timeMachine*60*60*1000);
  let yesterday = new Date(Date.now()-24*60*60*1000)
  let now = Utilities.formatDate(today, "GMT+8", "yyyy-MM-dd HH:mm") // HH 24HR
  let now_hr = Number(Utilities.formatDate(today, "GMT+8", "HH"))
  let now_string = now.split(' ')
  let now_date = now_string[0]
  let now_hrANDmin = now_string[1].split(':')

  if( now_hr < 6 ){
    now = Utilities.formatDate(yesterday, "GMT+8", "yyyy-MM-dd HH:mm") // HH 24HR
    now_string = now.split(' ')
    now_date = now_string[0]
  }
  hr = Number( now_hrANDmin[0] - interval_var) // 當前幾點 // 進前一時區

  function interval( now_hr_var, num ){
    if ( now_hr < 6 ){
      return 18
    }
    if ( now_hr < 12 ){
      return 0
    }
    if ( now_hr < 18 ){
      return 6
    }
    if ( now_hr < 24 ){
      return 12
    }
    return false
  }
  return {date:now_date,hr:interval()}
}
imgFactory.prototype.combitTime = function (start_Hr) {
    let gap = this.imgProperty[this.propertyWantToDraw]["timeGap"];
    function hr_min_producer(sign){
      let string_hr = '';
      let string_min = '';
      let hr_min_array = [];
      for (let hr = start_Hr; hr<start_Hr+6; hr+=1){
        if( hr < 10 ){
          string_hr = '0' + String(hr)
        }else{
          string_hr = String(hr)
        }

        let min = 0;
        while (min<60){
          if( min < 10 ){
            string_min = '0' + String(min)
          }else{
            string_min = String(min)
          }
          hr_min_array.push(`${string_hr}${sign}${string_min}`)
          min+=gap
        }
      }
      return hr_min_array
    }

    let day;
    let timeFormat = this.imgProperty[this.propertyWantToDraw].timeFormat;
    let now_hr = Number(Utilities.formatDate(new Date(Date.now()), "GMT+8", "HH"))
    switch (timeFormat) {
      case "YYYYMMDD[00-23][0-5]0":
        if( now_hr < 6 ){
          day = Utilities.formatDate(new Date(Date.now()-24*60*60*1000), "GMT+8", "yyyyMMdd")
          folderDay = Utilities.formatDate(new Date(Date.now()-24*60*60*1000), "GMT+8", "yyyyMMdd")
        }else(
          day = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd")
        )
        
        this.array = hr_min_producer('');
        this.array.forEach((value, index, array)=>{
          array[index] = day+value
        }) 
        break;
      case "YYYY-MM-DD-[00-23]-[0-5]0":
        if( now_hr < 6 ){
          day = Utilities.formatDate(new Date(Date.now()-24*60*60*1000), "GMT+8", "yyyy-MM-dd")
          folderDay = Utilities.formatDate(new Date(Date.now()-24*60*60*1000), "GMT+8", "yyyyMMdd")
        }else(
          day = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd")
        )
        this.array = hr_min_producer('-');
        this.array.forEach((value, index, array)=>{
          array[index] = `${day}-${value}`
        }) 
      break;
      case "YYYY-MM-DD_[00-23]00":
        if( now_hr < 6 ){
          day = Utilities.formatDate(new Date(Date.now()-24*60*60*1000), "GMT+8", "yyyy-MM-dd")
          folderDay = Utilities.formatDate(new Date(Date.now()-24*60*60*1000), "GMT+8", "yyyyMMdd")
        }else(
          day = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd")
        )
        this.array = hr_min_producer('');
        this.array.forEach((value, index, array)=>{
          array[index] = `${day}_${value}`
        })
      break;
      case "YYYY-MM-DD_0000":
        day = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd")
        this.array = [`${day}_0000`]

      break;

      default:
        console.log("Sorry");
    }
  
}
function img (variable, web='CWB',specialSelectArray){
  let imgType = new imgFactory(variable) //指定變數
  let now_interval = imgType.timeInterval(6) // 每次6hr 呼叫時間分4份 0-6 6-12 12-18 18-24
  imgType.combitTime(now_interval.hr)
  var getCategoryObj = imgType.imgProperty[imgType.propertyWantToDraw].category;
  var getCategoryObjList = specialSelectArray? specialSelectArray : Object.keys(getCategoryObj);
  getCategoryObjList.forEach((j)=>{
    let url = getCategoryObj[j].img_url
    let imgFormat = imgType.imgProperty[imgType.propertyWantToDraw].imgFormat
    
    imgType.array.forEach((item)=>{
      let id = folderId(j)
      let total_url = totalUrl(web,{url:url,time:item,imgFormat:imgFormat}) //url + String(item) + imgFormat
      getImg(total_url,id)              
    })
  })
}

function totalUrl(web,parameter){
  const {url,time,imgFormat} = parameter
  let total_url;
  switch(web){
    case 'CWB':
      total_url = url + String(time) + imgFormat;
      break;
    case 'NCDR':
      let midUrl = `${String(time).substring(0, 6)}/windmap_${String(time)}`
      total_url = url + midUrl + imgFormat // 202012/windmap_202012300800.png
      break;
  }
  Logger.log(total_url)
  return total_url
}

function img_part5(parameter){
  const getCategoryList = parameter; // [ '探空', 'JMA的圖' ]
  let url = null
  const jsonForfolderId = getExcelData();
  getCategoryList.forEach(( key )=>{
    const dataEachCategory = part5Data[ key ]
    folderDay_yesterday = Utilities.formatDate(new Date(Date.now()-(timeMachine+24)*60*60*1000), "GMT+8", "yyyyMMdd")
    const folderId = jsonForfolderId[folderDay_yesterday][key];
    dataEachCategory.ids.forEach((id)=>{
      dataEachCategory.times.forEach((time)=>{
        if ( key === '探空' ){
          url = dataEachCategory.url_start + folderDay_yesterday.slice(2,8) + time + `_${ id }` + dataEachCategory.url_end
        }else if ( key === 'JMA天氣圖' ){
          url = dataEachCategory.url_start + id + folderDay_yesterday + time + dataEachCategory.url_end
        }
        Logger.log(url ,folderId)
        getImg(url,folderId)
      })
    })

  })
}


