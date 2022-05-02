// ------------------------------------------------------------------
//
// 外部ファイルを取り込む
//
// ------------------------------------------------------------------

var com = require('./common.js');



// ------------------------------------------------------------------
// 
// 配列をCSV出力する【サンプル】
// 
// ------------------------------------------------------------------

const { release } = require('process');
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
// Excelファイルのユーザー一覧のシート内に罫線、テキスト貼り付けがあってもRead/Writeができるかの確認用
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
// 
// テキストファイル読込のテスト
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

