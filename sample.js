// ------------------------------------------------------------------
//
// 外部ファイルを取り込む
//
// ------------------------------------------------------------------

var com = require('./common.js');

'use strict';


// ------------------------------------------------------------------
// テキストファイル処理
// 配列をCSVファイルに出力する【サンプル】
// 
// ------------------------------------------------------------------

const { release } = require('process');
const { EPROTOTYPE } = require('constants');
function test01_write(flg_bom) {
 
    try {

        var fs = require('fs');
        var fd;
        var file = __dirname + "/test01_write" + ".txt";

        var ary = [];

        ary.push((flg_bom == "bom") ? "\ufeff" : "");     // 引数にbomがあれば、先頭行にBOMをつける
        ary.push("駅名,果物,価格");
        ary.push("米原,バナナ,1000");
        ary.push("彦根,りんご,3000");

        fs.writeFileSync( file , "" );                   // 空のファイルを書き出す
        
        fd = fs.openSync( file, "w");                    // 空ファイルをオープン
        
        fs.writeSync(fd, ary.join("\n"), 0);             // 引数は書き込みを開始する場所、0でOK
        fs.closeSync(fd);

    } catch(err){
        console.log(err.name + ': ' + err.message);
        process.exit(-1);
    }
};




// ------------------------------------------------------------------
//
// 新規ファイルを作成して、プログラム内から固定文字を書き込む
// 
// ------------------------------------------------------------------

function test_create_text() {

    var fs = require('fs'), fd;

    var path = __dirname + "/";             　　　// パス
    var file = "test.csv";                        // ファイル（先）

    // 空ファイルを新規作成
    fs.writeFileSync(path + file, "");

    // 空ファイルをAppendモードでOpen
    var fd = fs.openSync(path + file, "a");    

    var arry = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];

    for (let icnt = 0; icnt < 10; icnt++) {

        let data = [];

        if (icnt == 0) {
            data[0] = "カラム1";
            data[1] = "カラム2";
        } else {
            data[0] =  arry[icnt].toString();
            data[1] = (arry[icnt] * 10).toString();
        }

        // ファイルに１レコードを書き込む
        fs.writeSync(fd, data.join(",") + "\n", 0);
    }
    fs.closeSync(fd);
}




// ------------------------------------------------------------------
// Excelファイル処理
// Excelファイルを読み込んで、特定のセルの値を読み込み、新規のテキストファイルに書き込む
// 
// ------------------------------------------------------------------
function test02_sheet_user(){

    var fs   = require('fs');
    var file = __dirname + "/ユーザー一覧の書出しテスト" + ".txt";

    var XLSX     = require('xlsx');
    var workbook = XLSX.readFile('_端末ID割付表.xlsx');
    var worksheet = workbook.Sheets["ユーザー一覧"];
    var msg = "";
    var data_cnt = 0;
    var data1 = "";
    var data2 = "";
    var ary = [];

    if ( worksheet == undefined ){

        msg = "読み込む対象のシートは存在しません";
        console.log(msg);
        return;
    }
        
    fs.writeFileSync( file , "" );                                              // 空のファイルを書き出す
    var fd = fs.openSync( file, "a");                                           // ファイルをオープン

    for ( let irow = 0; irow < 15; irow++ ) {

        ary = [];

        for(let icol = 0; icol < 15; icol++ ) {

            if (worksheet[XLSX.utils.encode_cell({c:icol, r:irow})] == undefined) {
                data2 = "";                                                      // セルの値が空白の場合
            } else {
                data1 = worksheet[XLSX.utils.encode_cell({c:icol, r:irow})].v;   // セルの値を取得
                data2 = (data1 + "").replace(/\r?\n/g,"");                       // 文字列の中の改行を削除　+ ""をつけるのは、セルの値が数値の場合エラーになるのを回避するため、文字型に変換している
            }
            ary.push(data2);
        }

        let wrt_data = ary.join() + "\n";                                       // \nは改行コード
        if ( irow == 0 ){ wrt_data = "\ufeff" + wrt_data; } 
        fs.writeSync(fd, wrt_data, 0);                                          // ファイル書き込み
    }
    fs.closeSync(fd);
};




// ------------------------------------------------------------------
// テキストファイル処理
// テキストファイルを読み込んでコンソールに表示する
// 
// ------------------------------------------------------------------

function test03_read() {
    
    var fs   = require('fs');
    var fd;
    var filePath = __dirname + "/ユーザー一覧_User" + ".txt";
    
    var text  = fs.readFileSync(filePath, 'utf8');                 // 第二引数はテキストファイルの文字コードを指定
    var lines = text.toString().split('\n');                       // \r\n or \n 改行コード split 文字列⇒配列
    
    for (var idx in lines) {
        console.log(lines[idx]);
    }
};



// ------------------------------------------------------------------
// 連想配列
// 連想配列の値からキーを取得  結果がなければnullが返る
// ------------------------------------------------------------------

function test_write_sisetu01() {

    // Excelのシート名と楽楽の自動採番キーの１対１対応
    var hash_test = {
        "0001": "所沢市再発行",
        "0002": "NSKデモ",
        "0003": "野洲市",
        "0004": "美咲町",
        "0005": "守山市",
        "0006": "宗像市",
        "0007": "環境技研",
        "0008": "鉄道運輸機構(廃止)",
        "0009": "西和賀町",
        "0010": "東金市"
    };

    // 連想配列の値からキーを取得  結果がなければnullが返る
    var result = Object.keys(hash_test).reduce(function (r, k) { return hash_test[k] == "大阪市" ? k : r }, null);
    console.log(result);
}




// ------------------------------------------------------------------
// テキストファイル
// 端末ID.csvと請求.csvを読み込み、都道府県＋市町村名＋ユーザーキーでマッチングさせ、
// 結果を請求_new.csvに書き込む
// ------------------------------------------------------------------

function test_match_meisyo() {
 
    var fs = require('fs'), fd;

    var file_r = __dirname + "/端末ID/端末ID"    + ".csv";             // 読み込みファイル（元）
    var file_w = __dirname + "/端末ID/請求"      + ".csv";             // 書き込み先ファイル（先）
    var file_n = __dirname + "/端末ID/請求_new"  + ".csv";             // 書き込み先ファイル（先）  

    // 端末ID.csvを読み込む
    var text  = fs.readFileSync(file_r, 'utf8');                      // 第二引数はテキストファイルの文字コードを指定
    var lines = text.toString().split('\n');                          // \r\n or \n 改行コード split 文字列⇒配列
   
    var ary_user = {};                                         　     // 端末IDの連想配列（自動採番:ユーザー名）

    // 端末IDを全件読み込んで、連想配列に（自動採番：ユーザー名）をセット
    for (var idx in lines) {

        let ary_data = lines[idx].split(',')[0];
        if (ary_data != "" ) {
            if (ary_data != "自動採番") {
                ary_user[ary_data] = lines[idx].split(',')[1] + lines[idx].split(',')[2];    // ary_user[キー]=値 で連想配列に追加
            }
        }
    }

    // 請求.csvを読み込む
    var text  = fs.readFileSync(file_w, 'utf8');                  // 第二引数はテキストファイルの文字コードを指定
    var lines = text.toString().split('\n');                      // \r\n or \n 改行コード split 文字列⇒配列

    // 請求_new.csv　空ファイルを作成してオープン
    fs.writeFileSync(file_n, "");
    var fd = fs.openSync(file_n, "a");

    // 請求.csvを全件読み込む
    for (var idx in lines) {

        let wrt_lines = [];

        // replace 改行コードを削除
        // split 文字列⇒配列
        wrt_lines[0] = (lines[idx].replace(/\r/g, "")).split(',');
        wrt_lines[1] = "";
     

        if (idx == 0) {
            wrt_lines[1] = "自動採番（端末ID）";
        } else {

            //請求.csvの都道府県名＋市区町村名→端末ID.csvのユーザー名を検索してマッチングさせる
            let ary_user_key = Object.keys(ary_user).filter((key) => {

                return ary_user[key] === wrt_lines[0][1] + wrt_lines[0][2] + wrt_lines[0][3];
            })

            // 2回以上マッチングがあった場合、配列で複数返すため、先頭の1回目のみ出力対象とする            
            switch (ary_user_key.length) {
                case 0:
                    wrt_lines[1] = "";
                    break;
                case 1:
                    wrt_lines[1] = ary_user_key[0];
                    break;
                default:
                    wrt_lines[1] = "複数マッチ有 " + ary_user_key;
            }
        }
        
        // 請求_new.csvを書き込む
        // EOF行を書き込まない
        let wrt_lines_last = wrt_lines[0][0] + "," + wrt_lines[0][1] + wrt_lines[0][2] + "," + wrt_lines[1];
        if (wrt_lines[0][0] != "") { fs.writeSync(fd, wrt_lines_last + "\n", 0) };
    }
    fs.closeSync(fd);
}


// ------------------------------------------------------------------
// 日付の論理チェック
//
// https://blog.beatdjam.com/entry/2017/07/28/181201
//
// 画面に、20022/02/01形式で入力してもらい、日付の論理チェックをTorFで返す
// 2022/02/01でも、2022/2/1でも論理チェックOKとしている
// 2022-02-01型は×
//
//
// ------------------------------------------------------------------

function test_check_date(date) {
    
    if (date == undefined) {
        return false;
    }

    if (date == "") {
        return false;
    }
    
    // yyyy/mm/dd の形のみOKとする
    if (!date.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) {
        return false;
    }          

    var ddate = new Date(date);

    if (ddate.getFullYear() == date.split("/")[0]   && 
        ddate.getMonth()    == date.split("/")[1]-1 &&
        ddate.getDate()     == date.split("/")[2]) {
    } else {
        return false;
    }
    return true;
}



// ------------------------------------------------------------------
// 数値チェック
//
// 数値であるかどうかのチェック T or Fを返す
// 正規表現の説明 先頭が[0-9]のいずれかで始まり、終わりも[0-9]のいずれかで終わる
// +があるので[0-9]のいずれかが1個以上含まれる
//
// ------------------------------------------------------------------

function test_check_numeric(data) {
    
    var regex = new RegExp(/^[0-9]+$/);
    if (regex.test(data) == false) {
        return false;
    }
    return true;
}


// ------------------------------------------------------------------
// 日付処理
//
// システム日付を yyyy/mm/ddの形式で取得する
// 1/1 → 01/01 に変換する
// ------------------------------------------------------------------

function test_get_date() {

    let date = new Date();

    let yy = date.getFullYear();
    let mm = (date.getMonth() + 1).toString().padStart(2, '0');
    let dd = date.getDate().toString().padStart(2, '0');

    return yy + "/" + mm + "/" + dd;
}


// ------------------------------------------------------------------
// 日数の経過日を計算
//
// 引数の日付はyyyymm/dd型のちゃんとした日付でないといけない
// 1日＝86400000秒
// day2 < day1の時はマイナス日数となる
// day = day2 の時は0が返ってくる
//
// ------------------------------------------------------------------

function test_get_date_sa(data1, data2) {

    let day1 = new Date(data1);
    let day2 = new Date(data2);

    let day3 = (day2 - day1) / 86400000;
    return day3.toString();
}


// ------------------------------------------------------------------
//
// システム時刻を hh:mm:ddの形式で取得する
//
// ------------------------------------------------------------------

function test_get_time() {

    let date = new Date();

    let hh = date.getHours().toString().padStart(2, '0');
    let mm = date.getMinutes().toString().padStart(2, '0');
    let ss = date.getSeconds().toString().padStart(2, '0');

    return hh + ":" + mm + ":" + ss;
}


// ------------------------------------------------------------------
//
// 指定した月の月末日を取得する
// 11月の月末日を取得したい場合、Dateに11月0日を与えてやる、ちょっとしたテクニックがある
// 引数には、yyyy/mm形式でデータを与える
// まだ作成途中ではある
//
// ------------------------------------------------------------------

function test_get_date_lastdd(yy,mm) {

    let date = new Date(yy, mm, 0);
    let aa = date.getDate();

    return aa;
}



// ------------------------------------------------------------------
//
// DateのgetMonth()の使い方について、動作確認
//
// ------------------------------------------------------------------

function test_date1() {

   // 引数を""で囲まない場合、0が戻る
    var aa = new Date(2022/05/09);
    var bb = aa.getMonth();
    // console.log(bb);

    // 0からカウントされて、4が返る
    var aa = new Date("2022/05/09");
    var bb = aa.getMonth();
    // console.log(bb);

    // 0からカウントされて、4が返る
    // 1桁の場合、先頭0埋めしなくても問題ない
    var aa = new Date("2022/5/9");
    var bb = aa.getMonth();
    // console.log(bb);

    // aaは空白が返ってくる。yyyy/mm/ddであれば問題ないが、yyyy/mm形式は×
    // ccには2022 bbには4が返ってくる
    var aa = new Date("2022/5");
    var cc = aa.getFullYear();
    var bb = aa.getMonth();
    //console.log(aa);
    //console.log(cc);
    //console.log(bb);

    // 0からカウントされて、4が返る
    // /編集でなく、-編集でもOK
    var aa = new Date("2022-5-9");
    var bb = aa.getMonth();
    //console.log(bb);

    // 1からカウントされて、5が返る
    var aa = new Date("2022", "5", "9");
    var bb = aa.getMonth();
    // console.log(bb);


    // 5が返る
    // "05" でも"5"でもOK
    var aa = new Date("2022", "05", "9");
    var bb = aa.getMonth();
    //console.log(bb);


    // 5が返る
    // 文字列で渡しても、数値で渡してもどっちでもOK
    var aa = new Date(2022, 5, 9);
    var bb = aa.getMonth();
    //console.log(bb);

}


// ------------------------------------------------------------------
//
// DateのgetDay()の使い方について、動作確認
//
// ------------------------------------------------------------------

function test_date2() {


    // 0(日曜日)、 つまり、5月1日の曜日が返る
    var aa = new Date("2022/05/01");
    var bb = aa.getDay();
    console.log(bb);

    // 3(水曜日)、つまり、6月の曜日が返る 月は0からスタート 0が1月なので5月は6月となる 
    var aa = new Date(2022,5,1);
    var bb = aa.getDay();
    console.log(bb);

}
                   

// ------------------------------------------------------------------
//
// 日付の加算、減算の動作確認
// 
// ------------------------------------------------------------------

function test_date_add() {

    //システム日付から１年後
    let date = new Date();    
    date.setFullYear(date.getFullYear() + 1);
    let aa = date.getFullYear();
    console.log(aa);



    //指定した日から１ヶ月後
    let date1 = new Date("2000/12/01");
    date1.setMonth(date1.getMonth() + 2);
    let bb = date1.getFullYear().toString() + "/" + date1.getMonth().toString();
    console.log(bb);


    //指定した日から１日前
    let date2 = new Date("2000/01/01");
    date2.setDate(date2.getDate() - 1);
    let cc = date2.getFullYear().toString() + "/" + (date2.getMonth() + 1).toString() + "/" + date2.getDate().toString();
    // let cc = date2.getFullYear().toString();
    console.log(cc);



}



// ------------------------------------------------------------------
//
// switch文の中に条件式を組み込んだ使い方例
// 
// ------------------------------------------------------------------

function test_switch() {

    for (let cnt = 0; cnt <= 10; cnt++) {

        let ret;

        switch (true) {
            case cnt == 0:
                ret = "0 " + cnt;
                break;
            case [1, 2, 3].includes(cnt):
                ret = "1,2,3 " + cnt;
                break;
            case [4, 5, 6].includes(cnt):
                ret = "4,5,6 " + cnt;
                break;
            case (cnt >= 7 && cnt <= 8):
                ret = "7,8 " + cnt;
                break;
            default:
                ret = "9以上 " + cnt;
        }
        console.log("ret = " + ret);
    }
}




// ------------------------------------------------------------------
//
// switch文の中に条件式を組み込んだ使い方例
// 
// ------------------------------------------------------------------

function test_switch() {

    for (let cnt = 0; cnt <= 10; cnt++) {

        let ret;

        switch (true) {
            case cnt == 0:
                ret = "0 " + cnt;
                break;
            case [1, 2, 3].includes(cnt):
                ret = "1,2,3 " + cnt;
                break;
            case [4, 5, 6].includes(cnt):
                ret = "4,5,6 " + cnt;
                break;
            case (cnt >= 7 && cnt <= 8):
                ret = "7,8 " + cnt;
                break;
            default:
                ret = "9以上 " + cnt;
        }
        console.log("ret = " + ret);
    }
}

