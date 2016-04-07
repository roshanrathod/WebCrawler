var request = require('request');
var cheerio = require('cheerio');
var URL = require('url-parse');
var fs = require('fs');
var excelbuilder = require('msexcel-builder');

var UrlToParse = "http://www.google.com";
var NumOfPagesToVisit = 5;
var pagesVisited = {};
var numPagesVisited = 0;
var pagesToVisit = [];
var url = new URL(UrlToParse);
var baseUrl = url.protocol + "//" + url.hostname; 

pagesToVisit.push(UrlToParse);
crawl();

// Create a new workbook file in current working-path 
var workbook = excelbuilder.createWorkbook('./', 'file.xlsx')
  
// Create a new worksheet 
var sheet1 = workbook.createSheet('sheet1', 10, 500);

//Input tag details
sheet1.set(1,1,'Tag'); 
sheet1.set(2,1, 'Type');
sheet1.set(3,1, 'ID');
sheet1.set(4,1, 'Name');
sheet1.set(5,1, 'Value');
  
//start at row 2 of excel, since first is header
var rowNum = 2;
     
function crawl() {
  if(numPagesVisited >= NumOfPagesToVisit) {
    console.log("Reached limit of number of pages to visit.");
    return;
  }
  var nextPage = pagesToVisit.pop();
  console.log('Next:' + nextPage);
  if (nextPage in pagesVisited) {
    // We've already visited this page, so repeat the crawl
    crawl();
  } else {
    // New page we haven't visited
    visitPage(nextPage, crawl);
  }
}

function visitPage(url, callback) {
  // Add page to our set
  pagesVisited[url] = true;
  numPagesVisited++;

  // Make the request
  console.log("Visiting page " + url);
  request(url, function(error, response, body) {
     // Check status code (200 is HTTP OK)
     console.log("Status code: " + response.statusCode);
     if(response.statusCode !== 200) {
       callback();
       return;
     }
     // Parse the document body
     var $ = cheerio.load(body);
	 // fs.appendFileSync('body1.txt',body)
	 saveExcel($,url);
	 collectInternalLinks($);
	 callback();
     
  });
}

function collectInternalLinks($) {
	//fetch all anchor links from the page
    var relativeLinks = $("a");
    
	//Iterate and get href's for all and add to pagesToVisit
    relativeLinks.each(function() {
        pagesToVisit.push(baseUrl + $(this).attr('href'));
    });
}

function saveExcel($,url){
//add Page URL and blank lines
sheet1.set(1,rowNum,' ');
rowNum = rowNum+1;
sheet1.set(1,rowNum,url);
rowNum = rowNum+1;
sheet1.set(1,rowNum,' ');
rowNum = rowNum+1;

  $(':input').each(function( index ) {
	 	 
  var input = $(this);    
  sheet1.set(1,rowNum,input);
  sheet1.set(2,rowNum, input.attr('type'));
  sheet1.set(3,rowNum, input.attr('id'));
  sheet1.set(4,rowNum, input.attr('name'));
  sheet1.set(5,rowNum, input.val());
  
  //Next row
  rowNum = rowNum+1;
	
	
  });
// Save it 
  workbook.save(function(ok){
    if (!ok) 
      workbook.cancel();
    else
      console.log('Excel workbood created');
  });



}