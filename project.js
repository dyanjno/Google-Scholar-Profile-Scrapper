// npm init -y
// npm install minimist
// npm install axios
// npm install jsdom
// npm install excel4node
// npm install request
 
// node project.js --source= URL
 
let minimist = require("minimist");
let axios = require("axios");
let jsdom = require("jsdom");
let excel4node = require("excel4node");
let fs = require('fs');
let request = require('request');
 
 
let args = minimist(process.argv);
 
// Easy input using minimist
// download using axios
// read using jsdom
// make excel using excel4node
 
let responsePromise = axios.get(args.source);
responsePromise.then(function(response){
    let html = response.data;
    let dom = new jsdom.JSDOM(html);
    let document = dom.window.document;
 
    let name = document.querySelector('#gsc_prf_in').textContent;
    let disc = document.querySelector('.gsc_prf_il').textContent;
    let discarr = disc.split(",");
    let imageurl = document.querySelector('#gsc_prf_pua img').getAttribute('src');
   
    let spec = document.querySelectorAll('#gsc_prf_int a');
    let specs = [];
    for(let i = 0; i < spec.length; i++){
        specs.push(spec[i].textContent);
    }
   
    let titles = document.querySelectorAll('#gsc_a_b a.gsc_a_at');
    let titleName = [];
    for(let i = 0; i < titles.length; i++){
        titleName.push(titles[i].textContent);
    }
 
    let authors = document.querySelectorAll('#gsc_a_b div.gs_gray');
    let authorsName = [];
    for(let i = 0; i < authors.length; i++){
        authorsName.push(authors[i].textContent);
    }
   
    let citations = document.getElementsByClassName('gsc_a_ac gs_ibl');
    let citation = [];
    for(let i = 0; i < citations.length; i++){
        citation.push(citations[i].textContent);
    }
 
    let years = document.getElementsByClassName('gsc_a_h gsc_a_hc gs_ibl');
    let year = [];
    for(let i = 0; i < years.length; i++){
        year.push(years[i].textContent);
    }
 
    let coauthors = document.querySelectorAll('div.gsc_rsb_aa');
    let coauthor = [];
 
    for(let i = 0; i < coauthors.length; i++){
        let coauth = {
        };
 
        coauth.name = coauthors[i].querySelector('a').textContent;
        coauth.dis = coauthors[i].querySelector('.gsc_rsb_a_ext').textContent;
        coauthor.push(coauth);
    }
 
    // Creating Excel File
 
    let wb = new excel4node.Workbook();
 
    var myStyle = wb.createStyle({
        font: {
          size: 14,
          bold: true,
          underline: true,
        },
        alignment: {
          wrapText: true,
          horizontal: 'center',
        },
    });
 
    var normalStyle = wb.createStyle({
        font: {
          bold: true,
          underline: true,
        },
        alignment: {
          wrapText: true,
        },
    });
 
    var subStyle = wb.createStyle({
        font: {
          underline: true,
        },
        alignment: {
          wrapText: true,
          horizontal: 'center',
        },
    });
 
    // Sheet - 1
    let sheet = wb.addWorksheet("Profile");
 
    download(imageurl, name + '.png', function(){
        // Work
    });

    console.log(name + '.png');
    sheet.addImage({
        path: name + '.png',
        type: 'picture',
        position: {
            type: 'twoCellAnchor',
            from: {
                col: 5,
                colOff: 0,
                row: 5,
                rowOff: 0
            },
            to: {
                col: 7,
                colOff: 0,
                row: 10,
                rowOff: 0
            }
        }
    });
 
    sheet.cell(1,1).string(name).style(myStyle);
    sheet.column(1).setWidth(50);
    for(let i = 0; i < discarr.length; i++){
        sheet.cell(2 + i, 1).string(discarr[i]);
    }
    sheet.column(2).setWidth(1);
    sheet.column(3).setWidth(50);
    if(specs.length > 0){
        sheet.cell(1, 3).string("Specializations").style(myStyle);
        for(let i = 0; i < specs.length; i++){
            sheet.cell(2 + i, 3).string(specs[i]).style(subStyle);
        }
    }
 
    sheet.cell(5 + discarr.length, 1).string("Google Scholar Link :").style(normalStyle);
    sheet.cell(6 + discarr.length, 1).string(args.source);
 
    // Sheet - 2
 
    let sheet2 = wb.addWorksheet('Articles');
    sheet2work(sheet2, titleName, authorsName, citation, year, myStyle);
 
    // Sheet - 3
    let sheet3 = wb.addWorksheet('Co-Authors');
    sheet3work(sheet3, coauthor, myStyle);
 
    wb.write(name + "'s Profile.csv");
 
    // Sheet - 4
    sheet4work(wb, document);
   
}).catch(function(err){
    console.log(err);
})
 
 
function sheet2work(sheet2, titleName, authorsName, citation, year, myStyle){
    sheet2.cell(1,2).string("TOP CITED ARTICLES").style(myStyle);
    sheet2.column(1).setWidth(7);
    sheet2.column(2).setWidth(100);
    sheet2.cell(3,1).string("S. No.").style(myStyle);
    sheet2.cell(3,2).string("TITLE").style(myStyle);
    sheet2.cell(3,3).string("CITED BY").style(myStyle);
    sheet2.cell(3,4).string("YEAR").style(myStyle);
 
    for(let i = 0; i < titleName.length; i++){
        sheet2.cell(4 + 4 * i, 1).string(i + 1 + ".");
        sheet2.cell(4 + 4 * i, 2).string(titleName[i]);
        sheet2.cell(5 + 4 * i, 2).string(authorsName[2 * i]);
        sheet2.cell(6 + 4 * i, 2).string(authorsName[2 * i + 1]);
 
        sheet2.cell(4 + 4 * i, 3).string(citation[i]);
        sheet2.cell(4 + 4 * i, 4).string(year[i]);
    }
}

var download = function(uri, filename, callback){
    request.head(uri, function(err, res, body){
        request(uri).pipe(fs.createWriteStream(filename)).on('close', callback);
    });
};
 
function sheet3work(sheet3, coauthor, myStyle){
    sheet3.column(2).setWidth(30);
    sheet3.column(3).setWidth(90);
 
    sheet3.cell(1,2).string("CO-AUTHORS").style(myStyle);
    sheet3.cell(3,1).string("S. No.").style(myStyle);
    sheet3.cell(3,2).string("Name").style(myStyle);
    sheet3.cell(3,3).string("Description").style(myStyle);
   
    for(let i = 0; i < coauthor.length; i++){
        sheet3.cell(4 + i, 1).string(i + 1  + ".");
        sheet3.cell(4 + i, 2).string(coauthor[i].name);
        sheet3.cell(4 + i, 3).string(coauthor[i].dis);
    }
}
 
function sheet4work(wb, document){
    var myStyle = wb.createStyle({
        font: {
          size: 14,
          bold: true,
          underline: true,
        },
        alignment: {
          wrapText: true,
          horizontal: 'center',
        },
    });
 
    var subStyle = wb.createStyle({
        font: {
          size: 12,
          bold: true,
          underline: true,
        },
    });
 
    let sheet4 = wb.addWorksheet('Citations');
    sheet4.column(4).setWidth(20);
 
    sheet4.cell(1,4).string("CITATIONS").style(myStyle);
    sheet4.cell(4,1).string("CITATIONS").style(subStyle);
    sheet4.cell(5,1).string("h-index").style(subStyle);
    sheet4.cell(6,1).string("i10-index").style(subStyle);
    sheet4.cell(3,2).string("All").style(subStyle);
    sheet4.cell(3,3).string("Since 2017").style(subStyle);
   
    vals = document.querySelectorAll('.gsc_rsb_std');
 
    sheet4.cell(4,2).string(vals[0].textContent);
    sheet4.cell(4,3).string(vals[1].textContent);
    sheet4.cell(5,2).string(vals[2].textContent);
    sheet4.cell(5,3).string(vals[3].textContent);
    sheet4.cell(6,2).string(vals[4].textContent);
    sheet4.cell(6,3).string(vals[5].textContent);
 
    year = document.querySelectorAll('.gsc_g_t');
    cited = document.querySelectorAll('.gsc_g_al');
 
    sheet4.cell(3,5).string("YEAR").style(subStyle);
    sheet4.cell(3,6).string("CITED").style(subStyle);
 
    for(let i = 0 ; i < year.length; i++){
        sheet4.cell(4 + i, 5).string(year[i].textContent);
        sheet4.cell(4 + i, 6).string(cited[i].textContent);
    }
}