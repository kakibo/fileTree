'use strict';
const fs = require('fs');
const path = require('path');
const url = require('url');
const program = require('commander');
const winston = require('winston');
const Excel = require('exceljs');
const jschardet = require('jschardet');
const jconv = require('jconv');
const iconv = require('iconv-lite');

// Command Line Args
program
    .usage('[options] <Target Directory> [output excel path & name]')
    .option('-p, --path [url|pc|pcfull]', 'パスの表示方法', /^(url|pc|pcfull)$/i, 'url')
    .option('-f, --filter ["ext|ext"]', '拡張子フィルタ ※ダブルクォーテーションで囲み、半角「|」区切りで拡張子を指定してください。')
    .option('-s, --search ["正規表現"]', '検索文字列を指定。正規表現が利用できます。※ダブルクォーテーションで囲んでください。')
    .option('-T, --test', 'デバッグモード(詳細ログを出力）')
    .parse(process.argv);


// Logger - error: 0, warn: 1, info: 2, verbose: 3, debug: 4, silly: 5
var logger = new (winston.Logger)({
    transports: [
        new (winston.transports.Console)({ name: 'console', colorize: true, handleExceptions: true }),
        new (winston.transports.File)({
            name: 'file',
            filename: path.join('fileTree.log'),
            json: false,
            timestamp: function() { return (new Date()).toLocaleString() ; },
            handleExceptions: true
            })
    ],
    exitOnError: false
});
logger.log('info','==============================\nStart\n==============================');

const debug = program.test !== undefined ? true : false;
if(debug){
    logger.transports.console.level = 'debug';
    logger.transports.file.level = 'debug';
    logger.log('debug','-- debug mode --');
}

const debugLength = false;

//Settings
const opPath = program.path;
const opPathSep = opPath === 'url' ? '/' : '\\'
const opFilter = program.filter;
const opSearch = program.search;
const dir = program.args[0] || __dirname;
const outputExcel = program.args[1] || 'output.xlsx';

logger.log('debug','Target Dir : ' + dir);
logger.log('debug','Output Excel : ' + outputExcel);
logger.log('debug','opPath : ' + opPath);
logger.log('debug','Seperrator : ' + opPathSep);
logger.log('debug','Filter : ' + opFilter);

const reFilter = opFilter ? new RegExp('.*\.(' + opFilter + ')$') : undefined;
const reSearch = opSearch ? new RegExp(opSearch) : undefined;

const dirFull = path.resolve(dir);
const dirBaseName = path.basename(dir);
//const dirParent  = path.dirname(dir); //dirFull.slice(0, dirFull.indexOf(dir) + 1);
//logger.log('debug','Parent Dir : ' + dirParent);
const dirDept = path.dirname(dirFull).split(path.sep).length;


// Logic
let excel = (res)=>{
    logger.log('info','generate excel : ' + res.length +' record');

    let colLength = 0, pathLength = 0, titleLength = 0, remarkLength = 0;
    res.forEach( (item) =>{
        colLength = colLength > item.length ? colLength : item.length;
        pathLength = pathLength > item[0].length ? pathLength : item[0].length;
        titleLength = titleLength > item[1].length ? titleLength : item[1].length;
        remarkLength = remarkLength > item[2].length ? remarkLength : item[2].length;
    });
    pathLength = pathLength > 80 ? 80 : pathLength;
    titleLength = titleLength > 40 ? 80 : titleLength*2;
    remarkLength = remarkLength > 40 ? 80 : remarkLength*2;
    logger.log('debug', 'colLength: ' + colLength);
    logger.log('debug', 'pathLength: ' + pathLength);
    logger.log('debug', 'titleLength: ' + titleLength);
    
    let workbook = new Excel.Workbook();
    let worksheet = workbook.addWorksheet('Tree', {
            pageSetup:{fitToPage: true, fitToHeight:100, fitToWidth:1},
            views:[{state:'frozen', xSplit:1, ySplit:1}]
        });
    worksheet.spliceRows(1,'','');
    const colStyle1 = {
        border: {top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}},
    };
    const colStyle2 = {
        border: {top: {style:'thin'}, left: {style:'dotted'}, bottom: {style:'thin'}, right: {style:'dotted'}},
    };
    const colStyle3 = {
        border: {top: {style:'thin'}, left: {style:'dotted'}, bottom: {style:'thin'}, right: {style:'thin'}},
    };
    let headerColumns = [
        { header: 'パス', key: 'path', width: pathLength, style: colStyle1 },
        { header: 'タイトル', key: 'title', width: titleLength, style: colStyle1 },
        { header: '備考', key: 'remark', width: remarkLength, style: colStyle1 },
        { header: 'LV1', key: 'lv1', width: 5, style: colStyle2 }
    ];
    for(let i=2; i<colLength-2; i++){
        logger.log('debug', i);
        if(i === colLength-3){
            headerColumns.push( { header: 'LV'+i, key: 'lv'+i, width: 40, style: colStyle3 } );
        }else{
            headerColumns.push( { header: 'LV'+i, key: 'lv'+i, width: 5, style: colStyle2 } );
        }
    }
    worksheet.columns = headerColumns;
    worksheet.getRow(1).eachCell(function(cell, colNumber) {
        cell.fill = {type: 'pattern', pattern:'solid', fgColor:{argb:'ff999999'} };
    });
    
    res.forEach( (item) =>{
        worksheet.addRow(item);
    });

    workbook.xlsx.writeFile(outputExcel).then(()=>{
        logger.log('ifo', 'write ok!');
    });
};

var walk = (p, callback) => {
    let results = [];
    fs.readdir(p, (err, files) => {
        if (err) throw err;

        //console.log(files.length);
        let end = files.length;
        if(debugLength){console.log(end);}

        if (!end) {
            if(debugLength){console.log('end ' + end);}
            return callback(err, results);
        }
        
        files.forEach((item) => {
            let fp = path.resolve(path.join(p, item));
            let rp = fp.replace(dirFull,'');
            let bn = path.basename(item);
            let dp = path.dirname(fp).split(path.sep).length-dirDept-1;
            let ar = opPath === 'url' ? [url.parse(rp).pathname] : opPath === 'pc' ? [rp] : [fp];

            if(fs.statSync(fp).isDirectory()) {
                ar.push(''); // titile
                ar.push(''); // search
                for(let i=0; i<dp; i++){ar.push('');}
                let dirBn = bn + opPathSep;
                ar.push(dirBn);
                results.push(ar);
                
                walk(fp, (err, res)=>{
                    results = results.concat(res);
                    if (!--end) {
                        if(debugLength){console.log('dir end ' + end);}
                        callback(err, results);
                    }
                });
            }else{
                let checkTitle = (callback)=>{
                    const targetExt = opSearch ? /html|htm|txt|js|css/ : /html|htm/;
                    let searchRes, searchResLength = 0, searchResTxt = '';
                    if(targetExt.test(path.extname(item))){
                        logger.log('debug', path.extname(item) + ':' + bn);
                        fs.readFile(fp, (err, data)=>{
                            if (err) throw err;
                            let chardet = jschardet.detect(data),
                                text = data.toString();
                            logger.log('debug', chardet.encoding);
                            if(chardet.encoding !== 'UTF-8'){
                                if( /SHIFT_JIS|ISO-2022-JP|EUC-JP/.test(chardet.encoding) ){
                                    text = jconv.convert(data, chardet.encoding, 'UTF8').toString();
                                }else{
                                    text = iconv.decode(data, 'UTF-8').toString();
                                }
                            }
                            let start = text.indexOf('<title>');
                            let end = text.indexOf('</title>');
                            if(start > -1){
                                let title = text.slice(start+7, end);
                                ar.push(title);
                                logger.log('debug', title);
                            }
                            if(opSearch){
                                searchRes = text.match(reSearch);
                                if(searchRes){
                                    searchResLength = searchRes.length;
                                    searchResTxt = searchResLength +'件';
                                    searchRes.forEach((item)=>{
                                        searchResTxt += ', ' + item;
                                    });
                                    logger.log('debug', searchResTxt);
                                }
                            }
                            ar.push(searchResTxt);
                            callback();
                        });
                    }else{
                        ar.push(''); // title
                        ar.push(''); // search
                        callback();
                    }
                };
                checkTitle(()=>{
                    for(let i=0; i<dp; i++){ar.push('');}
                    ar.push(bn);
                    results.push(ar);
                    if (!--end) {
                        if(debugLength){console.log('file end ' + end);}
                        callback(err, results);
                    }
                });
            }
        });
    });
};

walk(dir, (err, results) => {
    if (err) new Error('get error',err);
    logger.log('debug','callback');
    results.sort((a,b)=>{
        var aa = a[0];
        var bb = b[0];
        if(aa < bb){return -1;}
        if(aa > bb){return 1;}
        return 0;
    });
    
    if(opFilter){
        results = results.filter((item)=>{
            return reFilter.test(item[0]);
        });
    }
    
    excel(results);
});

