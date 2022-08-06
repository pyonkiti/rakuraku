// ------------------------------------------------------------------
//
// 外部ファイルを取り込む
//
// ------------------------------------------------------------------

var com = require('./common.js');





// ------------------------------------------------------------------
// 
//  端末ID割付表.xlsxを読み込んで、全シートのヘッダ項目を、
//  端末ID管理表_タイトル抜粋.txtに書き出す
//
//  【メモ】
//     参考 https://photo-tea.com/p/17/excel-parse-javascript-node/
//     ['A1'].v 実データ.t データ型　w：表示されているデータ.vを外すとobjectが取得される
// 
// ------------------------------------------------------------------

function proc1_write_header() {

    var fs = require('fs');
    var fd;
    var file = __dirname + "/端末ID管理表_タイトル抜粋" + ".txt";

    var XLSX       = require('xlsx');
    var workbook   = XLSX.readFile('_端末ID割付表.xlsx');
    var sheet_list = workbook.SheetNames;                  // 全てのシートを取得

    var msg      = "";
    var data_cnt = 0;
    
    fs.writeFileSync( file , "" );                         // 空のファイルを書き出す
    fd = fs.openSync( file, "a");                          // ファイルをオープン
    
    msg = "全シート数 : " + sheet_list.length + " 件";
    console.log(msg);

    // 全シートを読み込む
    for (let i = 0; i < sheet_list.length; i++){
        
        // 読み飛ばし条件
        let ary_out = ["導入数集計", "ユーザー一覧", "_", "...", "..", "__", "-"];
        if ( ary_out.includes(sheet_list[i]) ) {

            msg = "読み飛ばし : " + sheet_list[i];
            console.log(msg);
            continue;
        }

        let Sheet1 = workbook.Sheets[sheet_list[i]];                                    // シート名を取得
        let wdata = [];
        let ary = [];

        data_cnt ++;
        ary.push(data_cnt.toString());                                                  // データ数カウント

        // 先頭列を読み込む
        for (let l = 0; l < 30; l++) {

            if (Sheet1[XLSX.utils.encode_cell({ c: l, r: 0 })] == undefined) {
                wdata[1] = "";                                                         // セルの値が空白の場合
            } else {
                wdata[0] = Sheet1[XLSX.utils.encode_cell({ c: l, r: 0 })].v;           // セルの値を取得 undefinedのセル（空白のセル）に.vするとエラーになる事象を回避している
                wdata[1] = (wdata[0] + "").replace(/\r?\n/g, "");                      // 文字列の中の改行を削除　+ ""をつけるのは、セルの値が数値の場合エラーになるのを回避するため、文字型に変換している
            }
            ary.push(wdata[1])
        }

        // ファイルの先頭にBOMをつけるとExcelで開いた時に文字化け回避 '\ufeff'がBOM
        wdata[2] = (data_cnt == 1) ? "\ufeff" + ary.join() + "\n" : ary.join() + "\n";

        fs.writeSync(fd, wdata[2], 0);                                                  // ファイル書き込み
    }
    fs.closeSync(fd);
   
    msg = "書込み件数 : " + data_cnt + " 件";
    console.log(msg); 
}





// ------------------------------------------------------------------
// 
// ユーザー一覧のユーザー名の一覧と、各シート名の一覧をテキスト出力する
//   作成：ユーザー一覧_User.txt
//   作成：ユーザー一覧_Shet.txt
//
// ------------------------------------------------------------------

function proc2_get_user() {

    var fs   = require('fs');
    var fd;
    var file_user = __dirname + "/ユーザー一覧_User" + ".txt";
    var file_shet = __dirname + "/ユーザー一覧_Shet" + ".txt";

    var XLSX       = require('xlsx');
    var workbook   = XLSX.readFile('_端末ID割付表.xlsx');
    var sheet_list = workbook.SheetNames;                        // 全てのシートを取得
    var sheet      = workbook.Sheets[sheet_list[0]];             // 先頭の１シート名を取得

    var wdata = [];
    var cnt   = 0;
    var msg   = "";

    fs.writeFileSync( file_user , "" );                          // 空のファイルを書き出す
    fd = fs.openSync( file_user, "a");                           // ファイルをオープン

    // ユーザー一覧のシートからユーザー名を取得
    for (let icol = 0; icol < 400; icol++) {

        // 読み飛ばし条件
        if ( sheet[XLSX.utils.encode_cell({c:1, r:icol})] == undefined ) {
                continue;
        } else {
            if ( sheet[XLSX.utils.encode_cell({c:1, r:icol})].v == "ユーザ名" ) {
                continue;
            }
        }
        
        cnt++;
        wdata = [];

        wdata[0] = sheet[XLSX.utils.encode_cell({ c: 1, r: icol })].v;
        wdata[1] = (wdata[0]).replace(/\r?\n/g,"");                                 // セルの中に改行があれば削除
        wdata[2] = wdata[1] + "\n";
        wdata[3] = (cnt == 1) ? "\ufeff" + wdata[2] : wdata[2];                     // 先頭行のみBOMを付加

        fs.writeSync(fd, wdata[3], 0);                                              // ファイル書き込み
    }
    fs.closeSync(fd);

    cnt = 0;
    fs.writeFileSync( file_shet , "" );                                            // 空のファイルを書き出す
    fd = fs.openSync( file_shet, "a");                                             // ファイルをオープン
    
    // 各シートからシート名を取得
    for ( let i = 0; i < sheet_list.length; i++ ) {
        
        // 読み飛ばし条件
        let ary_out = ["導入数集計", "ユーザー一覧", "_", "...", "..", "__", "-"];
        if ( ary_out.includes(sheet_list[i]) ) {
            
            msg = "読み飛ばし : " + sheet_list[i];
            console.log(msg);
            continue;
        }

        cnt++;
        wdata = [];

        wdata[0] = sheet_list[i];
        wdata[1] = wdata[0] + "\n";                                                // \nは改行コード
        wdata[2] = (cnt == 1) ? "\ufeff" + wdata[1] : wdata[1];                    // 先頭行のみBOMを付加

        fs.writeSync(fd, wdata[2], 0);                                             // ファイル書き込み
    }
    fs.closeSync(fd);
}





// ------------------------------------------------------------------
// 
// ユーザー一覧のユーザー名の一覧が書き込まれているテキストと、
// Excelのシート名の一覧が書き込まれているテキストとをマッチングさせて、
// User_newというテキストに書き込む
// マッチングは、シート名を元にして、ユーザー一覧にあるかを調べる
// また、マッチング結果を「マチング結果.txt」にログ出力する
// 
// ------------------------------------------------------------------

function proc3_match_user() {

    var fs = require('fs');                                                             // テキストファイルのモジュール
    var fd_user, fd_shet, fd_test;
    var fil_user = "/ユーザー一覧_User";
    var fil_shet = "/ユーザー一覧_Shet";
    var fil_test = "/マッチング結果";
    var txt_user = "", txt_shet = "";
    var ary_user = [], ary_shet = [];

    var wrt_user, wrt_logs;

    var match_cnt;                                                                      // 検索マッチフラグ
    var write_cnt = 0;                                                                  // 書き込みカウント（先頭にBOMをつけるため）

    txt_user = fs.readFileSync(__dirname.toString() + fil_user + ".txt", 'utf8');       // ファイルを読み込む（ReadOnly）
    txt_shet = fs.readFileSync(__dirname.toString() + fil_shet + ".txt", 'utf8');

    ary_user = txt_user.toString().split('\n');                                         // 読み込んだファイルを配列にセット
    ary_shet = txt_shet.toString().split('\n');

    fs.writeFileSync(__dirname.toString() + fil_user + "_new" + ".txt", "");            // 空ファイルを新規作成
    fs.writeFileSync(__dirname.toString() + fil_test + ".txt", "");

    fd_user = fs.openSync(__dirname.toString() + fil_user + "_new" + ".txt", "a");      // 書き込み先ファイルをオープン
    fd_test = fs.openSync(__dirname.toString() + fil_test + ".txt", "a");

    // マッチング結果のヘッダ更新
    wrt_logs = "";
    wrt_logs = wrt_logs + "usr";
    wrt_logs = wrt_logs + " : ";
    wrt_logs = wrt_logs + "sht";
    wrt_logs = wrt_logs + " : ";
    wrt_logs = wrt_logs + "ユーザー名";
    wrt_logs = wrt_logs + " : ";
    wrt_logs = wrt_logs + "シート名";
    wrt_logs = wrt_logs + "\n";

    fs.writeSync(fd_test, wrt_logs, 0);                                                 // ファイルへの書き込み（ヘッダのみ）

    // ユーザー名ファイルの配列をループ
    for (let idx_u in ary_user) {

        match_cnt = 0;
        wrt_logs = "";
        wrt_user = ary_user[idx_u];

        // シート名ファイルの配列をループ
        for (let idx_s in ary_shet) {

            // マッチしたデータを取得　（ユーザー一覧にシート名が含まれているかでマッチング）
            if (ary_shet[idx_s] !== "") {

                // 先頭行にBOMがあれば削除
                if (ary_shet[0].includes("\ufeff")) {
                    ary_shet[0] = ary_shet[0].replace("\ufeff", "")
                }

                if (ary_user[idx_u].includes(ary_shet[idx_s])) {

                    wrt_logs = wrt_logs + (" ".repeat(3) + idx_u.toString()).slice(-3);
                    wrt_logs = wrt_logs + " : ";
                    wrt_logs = wrt_logs + (" ".repeat(3) + idx_s.toString()).slice(-3);
                    wrt_logs = wrt_logs + " : ";
                    wrt_logs = wrt_logs + ary_user[idx_u];
                    wrt_logs = wrt_logs + " : ";
                    wrt_logs = wrt_logs + ary_shet[idx_s];
                    wrt_logs = wrt_logs + "\n";

                    wrt_user = wrt_user + "," + ary_shet[idx_s] + "\n";

                    match_cnt++;
                }
            }
        }

        // マッチしなかった場合
        if (match_cnt == 0) {

            wrt_logs = wrt_logs + (" ".repeat(3) + idx_u.toString()).slice(-3);
            wrt_logs = wrt_logs + " : ";
            wrt_logs = wrt_logs + " ".repeat(3);
            wrt_logs = wrt_logs + " : ";
            wrt_logs = wrt_logs + ary_user[idx_u];
            wrt_logs = wrt_logs + " : ";
            wrt_logs = wrt_logs + "マッチングなし";
            wrt_logs = wrt_logs + "\n";
        }

        // 最終行は読み飛ばす
        if (ary_user[idx_u] !== "") {
            if (write_cnt == 0) {
                wrt_user = "\ufeff" + wrt_user;
                wrt_logs = "\ufeff" + wrt_logs;
            }
            fs.writeSync(fd_user, wrt_user, 0);                                          // _newに書き込み
            fs.writeSync(fd_test, wrt_logs, 0);                                          // マッチした結果を出力

            write_cnt++;
        }
    }
    fs.closeSync(fd_user);                                                               // 書き込みファイルのクローズ
    fs.closeSync(fd_test);
}







// ------------------------------------------------------------------
//
// ユーザー一覧_Shetのテキストファイルに、マッチしたデータにチェックをつける
//
// ------------------------------------------------------------------
function proc4_match_shet() {

    var fs = require('fs');
    var fd_sheN;

    var fil_user = __dirname + "/ユーザー一覧_User_new" + ".txt";
    var fil_shet = __dirname + "/ユーザー一覧_Shet" + ".txt";
    var fil_sheN = __dirname + "/ユーザー一覧_Shet_new" + ".txt";

    var txt_user = fs.readFileSync(fil_user, 'utf8');        // 第二引数はテキストファイルの文字コードを指定
    var txt_shet = fs.readFileSync(fil_shet, 'utf8');

    var lin_user = txt_user.toString().split('\n');          // \r\n or \n 改行コード split 文字列⇒配列
    var lin_shet = txt_shet.toString().split('\n');

    var wrt_sheN;
    var match_cnt;

    fs.writeFileSync(fil_sheN, "");                         // 書き込みファイルの空ファイルを作成
    fd_sheN = fs.openSync(fil_sheN, "a");                   // 書き込み先ファイルをオープン

    // シート名一覧ファイルをループ
    for (let idx_shet in lin_shet) {

        match_cnt = 0;
        wrt_sheN = "";

        // ユーザー名一覧ファイルをループ
        for (let idx_user in lin_user) {

            // シート名とユーザー名が同じか確認
            if (lin_shet[idx_shet] == lin_user[idx_user].split(",")[1]) {

                match_cnt++;
                // シート名一覧ファイルに「match」を書き込む
                fs.writeSync(fd_sheN, lin_shet[idx_shet] + ",match," + match_cnt.toString() + "\n", 0);
            }
        }

        if (match_cnt == 0) {

            // マッチしなければ、シート名一覧ファイルに、シート名だけを書き込む
            fs.writeSync(fd_sheN, lin_shet[idx_shet] + ",," + "\n", 0);
        }
    }
    fs.closeSync(fd_sheN);                                  // 書き込みファイルのクローズ
}







// ------------------------------------------------------------------
//
// 端末ID割付表の全施設のシートをCSV出力する
//
// ------------------------------------------------------------------

function proc5_write_sisetu() {

    // テキストファイル
    var fs = require('fs'), fd;
    var file_sht = __dirname + "/施設テーブル" + ".csv";

    // Excel
    var xlsx = require('xlsx');
    var book = xlsx.readFile('_端末ID割付表_0411.xlsx');
    var sheets = book.SheetNames;                             // 全てのシートを取得

    var head_sheet = ["自動採番（ユーザー）", "シート名", "施設番号", "施設名", "SIM番号", "FOMA回線番号", "認証ID", "認証パスワード", "固定IP", "契約プラン", "客出しプラン", "接続タイプ", "現況", "日報", "リアルタイムトレンド", "ご利用開始日", "廃止日", "備考", "導入年度（期）"];
    var head = ["自動採番（ユーザー）", "施設番号", "施設名", "SIM番号", "FOMA回線番号", "認証ID", "認証パスワード", "固定IP", "契約プラン", "客出しプラン", "接続タイプ", "現況", "日報", "リアルタイムトレンド", "ご利用開始日", "廃止日", "備考", "導入年度（期）"];
    var wdata = [];
    var cnt = 0;
    var msg = "";

    fs.writeFileSync(file_sht, "");                  　       // 空のファイルを書き出す
    fd = fs.openSync(file_sht, "a");                          // ファイルをオープン

    // 先頭行のみBOMをつける
    // ヘッダ名称の１行だけを更新
    fs.writeSync(fd, "\ufeff" + head.join(",") + "\n", 0);    // ファイル書き込み
    fs.closeSync(fd);

    fd = fs.openSync(file_sht, "a");                          // ファイルをオープン

    // 全シートを順に読み込む
    for (let i = 0; i < sheets.length; i++) {

        // シートを読み飛ばす
        let ary_out = ["所沢市", "ユーザー一覧 (移行用)", "ユーザー一覧 (移行用) (後で消す)", "導入数集計", "ユーザー一覧", "_", "...", "..", "__", "-"];
        if (ary_out.includes(sheets[i])) { continue; }

        // 連想配列より、値からキーを取得  結果がなければnullが返る
        let hash_key = Object.keys(com.hash).reduce(function (r, k) { return com.hash[k] == sheets[i] ? k : r }, null);

        // ハッシュに値に該当するキーがない場合
        if (hash_key == null) {
            console.log("連想配列に存在しない値です : %s ", sheets[i]);
            continue;
        }

        cnt++;
        wdata = [];

        // 楽楽の自動採番キー、シート名を取得
        // シート名は確認チェック用のため、楽楽への本番移行では不要
        // wdata[1] = hash_key + "," + sheets[i];
        wdata[1] = hash_key;

        // 先頭行のみ、BOMとヘッダ名称を付加
        // wdata[1] = (cnt == 1) ? "\ufeff" + wdata[0] : wdata[0];

        // １シートの全データを取得
        var sheet = book.Sheets[sheets[i]];

        // １シート内の情報を取得
        // 1000回ループさせているのは、1000列以上データがない筈というMAX
        for (let irow = 0; irow < 1000; irow++) {

            // 読み飛ばし条件  c:列 r:行
            // 施設番号＋施設名＋SIM番号 共に空白なら読み飛ばす
            // 施設番号＋施設名＋SIM番号は共に空白だが、通番だけ入っている行がある　これは読み飛ばすようにする
            let flg_brk = 0;
            for (let icol = 1; icol < 4; icol++) {
                if (sheet[xlsx.utils.encode_cell({ c: icol, r: irow })] == undefined) {
                    flg_brk++;
                }
            }

            // １列に対し、３カラムが全て空白であれば、その行は飛ばす
            // 空白行を挟んだ後、再び施設データがセットされる場合があるため、breakせずに、continueする（1000回までループさせる）
            if (flg_brk == 3) { continue; }

            // ヘッダの１行目を読み飛ばす条件
            let ary_out_h = ["153574", 153574, "通番"];
            if (ary_out_h.includes(com.null_pad(sheet[xlsx.utils.encode_cell({ c: 0, r: irow })]))) { continue; }

            wdata[2] = ""
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 1, r: irow })]);        // 施設番号
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 2, r: irow })]);        // 施設名
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 3, r: irow })]);        // SIM番号
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 4, r: irow })]);        // FOMA回線番号
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 5, r: irow })]);        // 認証ID
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 6, r: irow })]);        // 認証パスワード
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 7, r: irow })]);        // 固定IP
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 8, r: irow })]);        // 契約プラン
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 9, r: irow })]);        // 客出しプラン
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 10, r: irow })]);       // 接続タイプ
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 11, r: irow })]);       // 現況
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 12, r: irow })]);       // 日報
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 13, r: irow })]);       // リアルタイムトレンド
            wdata[2] = wdata[2] + "," + com.get_date(sheet[xlsx.utils.encode_cell({ c: 14, r: irow })]);       // ご利用開始日
            wdata[2] = wdata[2] + "," + com.get_date(sheet[xlsx.utils.encode_cell({ c: 15, r: irow })]);       // 廃止日
            // wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 16, r: irow })]);    // ルータシリアル（移行不要）

            //17列目は、コルソスシリアル(通報) と備考の２種類ある （コルソスシリアル(通報) は移行不要）
            if (com.null_pad(sheet[xlsx.utils.encode_cell({ c: 17, r: 0 })]) == "備考") {
                wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 17, r: irow })]);
            } else {
                wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 17 + 2, r: irow })]);
            }

            // 18行目は、導入年度（期）とコルソスシリアル(電源)の２種類ある　（コルソスシリアル(電源)は移行不要）
            if (com.null_pad(sheet[xlsx.utils.encode_cell({ c: 18, r: 0 })]) == "導入年度（期）") {
                wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 18, r: irow })]);
            } else {
                wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 18 + 2, r: irow })]);
            }

            // 改行があれば削除
            // 改行が複数あっても全ての改行を削除してくれている
            wdata[3] = wdata[1] + wdata[2].replace(/\r?\n/g, "");

            fs.writeSync(fd, wdata[3] + "\n", 0);                                 // ファイル書き込み       
        }
    }
    fs.closeSync(fd);
}




// ------------------------------------------------------------------
//
// 端末ID割付表の特定の施設のシートだけをCSV出力する
// 楽楽の自動採番も決めうちで付加する
//
// ------------------------------------------------------------------

function proc6_write_sisetu_only() {

    // テキストファイル
    var fs = require('fs'), fd;
    var file_sht = __dirname + "/施設テーブル_追加分01" + ".csv";

    // Excel
    var xlsx = require('xlsx');
    var book = xlsx.readFile('_端末ID割付表_0411.xlsx');
    var sheets = book.SheetNames;                             // 全てのシートを取得

    var head_sheet = ["自動採番", "自動採番（ユーザー）", "シート名", "施設番号", "施設名", "SIM番号", "FOMA回線番号", "認証ID", "認証パスワード", "固定IP", "契約プラン", "客出しプラン", "接続タイプ", "現況", "日報", "リアルタイムトレンド", "ご利用開始日", "廃止日", "備考", "導入年度（期）"];
    var head = ["自動採番", "自動採番（ユーザー）", "施設番号", "施設名", "SIM番号", "FOMA回線番号", "認証ID", "認証パスワード", "固定IP", "契約プラン", "客出しプラン", "接続タイプ", "現況", "日報", "リアルタイムトレンド", "ご利用開始日", "廃止日", "備考", "導入年度（期）"];
    var wdata = [];
    var cnt = 0;
    var msg = "";
    var saiban_cnt = 2194 - 1;                                // 楽楽側の自動採番（開始番号 - 1）

    fs.writeFileSync(file_sht, "");                  　       // 空のファイルを書き出す
    fd = fs.openSync(file_sht, "a");                          // ファイルをオープン

    // 先頭行のみBOMをつける
    // ヘッダ名称の１行だけを更新
    fs.writeSync(fd, "\ufeff" + head.join(",") + "\n", 0);    // ファイル書き込み
    fs.closeSync(fd);

    fd = fs.openSync(file_sht, "a");                          // ファイルをオープン

    // 全シートを順に読み込む
    for (let i = 0; i < sheets.length; i++) {

        // 特定のシートのみ読み込む
        let ary_in = ["東かがわ市"];
        if (ary_in.includes(sheets[i]) == false) { continue; }

        // 連想配列より、値からキーを取得  結果がなければnullが返る
        let hash_key = Object.keys(com.hash).reduce(function (r, k) { return com.hash[k] == sheets[i] ? k : r }, null);

        // ハッシュに値に該当するキーがない場合
        if (hash_key == null) {
            console.log("連想配列に存在しない値です : %s ", sheets[i]);
            continue;
        }

        cnt++;        
        wdata = [];

        // 楽楽の自動採番キー、シート名を取得
        // シート名は確認チェック用のため、楽楽への本番移行では不要
        // wdata[1] = hash_key + "," + sheets[i];

        // 先頭行のみ、BOMとヘッダ名称を付加
        // wdata[1] = (cnt == 1) ? "\ufeff" + wdata[0] : wdata[0];

        // １シートの全データを取得
        var sheet = book.Sheets[sheets[i]];

        // １シート内の情報を取得
        // 1000回ループさせているのは、1000列以上データがない筈というMAX
        for (let irow = 0; irow < 1000; irow++) {

            // 読み飛ばし条件  c:列 r:行
            // 施設番号＋施設名＋SIM番号 共に空白なら読み飛ばす
            // 施設番号＋施設名＋SIM番号は共に空白だが、通番だけ入っている行がある　これは読み飛ばすようにする
            let flg_brk = 0;
            for (let icol = 1; icol < 4; icol++) {
                if (sheet[xlsx.utils.encode_cell({ c: icol, r: irow })] == undefined) {
                    flg_brk++;
                }
            }

            // １列に対し、３カラムが全て空白であれば、その行は飛ばす
            // 空白行を挟んだ後、再び施設データがセットされる場合があるため、breakせずに、continueする（1000回までループさせる）
            if (flg_brk == 3) { continue; }

            // ヘッダの１行目を読み飛ばす条件
            let ary_out_h = ["153574", 153574, "通番"];
            if (ary_out_h.includes(com.null_pad(sheet[xlsx.utils.encode_cell({ c: 0, r: irow })]))) { continue; }
            
            saiban_cnt++;

            // 自動採番と自動採番（ユーザー）を付加
            wdata[1] = saiban_cnt.toString().padStart(5, "0") + "," + hash_key;

            wdata[2] = ""
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 1, r: irow })]);        // 施設番号
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 2, r: irow })]);        // 施設名
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 3, r: irow })]);        // SIM番号
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 4, r: irow })]);        // FOMA回線番号
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 5, r: irow })]);        // 認証ID
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 6, r: irow })]);        // 認証パスワード
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 7, r: irow })]);        // 固定IP
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 8, r: irow })]);        // 契約プラン
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 9, r: irow })]);        // 客出しプラン
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 10, r: irow })]);       // 接続タイプ
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 11, r: irow })]);       // 現況
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 12, r: irow })]);       // 日報
            wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 13, r: irow })]);       // リアルタイムトレンド
            wdata[2] = wdata[2] + "," + com.get_date(sheet[xlsx.utils.encode_cell({ c: 14, r: irow })]);       // ご利用開始日
            wdata[2] = wdata[2] + "," + com.get_date(sheet[xlsx.utils.encode_cell({ c: 15, r: irow })]);       // 廃止日
            // wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 16, r: irow })]);    // ルータシリアル（移行不要）

            //17列目は、コルソスシリアル(通報) と備考の２種類ある （コルソスシリアル(通報) は移行不要）
            if (com.null_pad(sheet[xlsx.utils.encode_cell({ c: 17, r: 0 })]) == "備考") {
                wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 17, r: irow })]);
            } else {
                wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 17 + 2, r: irow })]);
            }

            // 18行目は、導入年度（期）とコルソスシリアル(電源)の２種類ある　（コルソスシリアル(電源)は移行不要）
            if (com.null_pad(sheet[xlsx.utils.encode_cell({ c: 18, r: 0 })]) == "導入年度（期）") {
                wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 18, r: irow })]);
            } else {
                wdata[2] = wdata[2] + "," + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 18 + 2, r: irow })]);
            }

            // 改行があれば削除
            // 改行が複数あっても全ての改行を削除してくれている
            wdata[3] = wdata[1] + wdata[2].replace(/\r?\n/g, "");

            fs.writeSync(fd, wdata[3] + "\n", 0);                                 // ファイル書き込み       
        }
    }
    fs.closeSync(fd);
}




// ------------------------------------------------------------------
//
// 指定したテキストファイルを読み込み、指定した行に対し、先頭０埋めをする
//
// ------------------------------------------------------------------

function proc7_write_zeropad() {

    var fs = require('fs');
    var fd;
    var file_red = __dirname + "/障害DB/障害_エンドユーザーリンクコード_コンバート用_元" + ".csv";
    var file_wrt = __dirname + "/障害DB/障害_エンドユーザーリンクコード_コンバート用_先" + ".csv";

    // 読み込むファイルをオープンして、全件読み込んで配列にセットする
    // 第二引数はテキストファイルの文字コードを指定
    // \r\n or \n 改行コード split 文字列⇒配列
    var text = fs.readFileSync(file_red, 'utf8');
    var lines = text.toString().split('\n');

    // 書き込むファイルをオープンする
    fs.writeFileSync(file_wrt, "");                  　       // 空のファイルを作成
    fd = fs.openSync(file_wrt, "a");                          // ファイルをオープン

    for (var idx in lines) {

        let wdata1 = [];

        // １列に対し、文字列⇒配列
        let wdata2 = lines[idx].toString().split(",");

        // ヘッダ名称
        if (idx == 0) {

            wdata1[0] = wdata2[0];
            wdata1[1] = wdata2[1];

        } else {

            wdata1[0] = (wdata2[0] == undefined || wdata2[0] == "") ? "" : wdata2[0].toString().padStart(6, "0");
            wdata1[1] = (wdata2[1] == undefined || wdata2[1] == "") ? "" : wdata2[1].toString().padStart(4, "0");
        }

        //　これをしないと、読込元のEOF行が書き込まれる
        if (wdata1[0] != "") {
            fs.writeSync(fd, wdata1[0] + "," + wdata1[1] + "\n", 0);                                                       // ファイル書き込み
        }
    }
    fs.closeSync(fd);
}





// ------------------------------------------------------------------
//
// 障害情報.xlsxを読み込み、対応者を取得する。ハッシュ（社員コード：社員名）から、対応者（社員コード）に対応する社員名を取得する。
// 障害情報.xlsxの自動採番（ID）と社員名だけをテキストファイルに出力する
// 出力したテキストファイルは楽楽にインポートする
//
// ------------------------------------------------------------------

function proc8_write_syaincd() {

    // テキストファイル
    var fs = require('fs'), fd;
    var file_sht = __dirname + "/障害DB/障害情報_対応者_コンバート用" + ".csv";

    // Excel
    var xlsx = require('xlsx');
    var book = xlsx.readFile('./障害DB/障害情報.xlsx');
    var sheet = book.Sheets["障害情報"];                    // １シートの全データを取得

    var head = ["自動採番", "対応者（主）", "対応者コード"];
    var wdata = [];
    var msg = "";

    // Excelのシートの存在チェック
    if (sheet == undefined) {

        msg = "読み込む対象のシートは存在しません";
        console.log(msg);
        return;
    }

    fs.writeFileSync(file_sht, "");                  　       // 空のファイルを作成
    fd = fs.openSync(file_sht, "a");                          // ファイルをオープン

    // 先頭行のみBOMをつける
    // ヘッダ名称の１行だけを更新
    fs.writeSync(fd, "\ufeff" + head.join(",") + "\n", 0);    // ファイル書き込み
    fs.closeSync(fd);

    fd = fs.openSync(file_sht, "a");                          // ファイルをオープン

    wdata = [];

    // １シート内の情報を取得  
    for (let irow = 0; irow < 300; irow++) {

        // 読み飛ばし条件  c:列 r:行
        // ID行が空白になればbreak           
        if (sheet[xlsx.utils.encode_cell({ c: 0, r: irow })] == undefined) { break; }

        // ヘッダの１行目を読み飛ばす条件
        let ary_out_h = ["ID"];
        if (ary_out_h.includes(com.null_pad(sheet[xlsx.utils.encode_cell({ c: 0, r: irow })]))) { continue; }

        // Excelシートから対応者コードを取得
        let shain = [];
        shain[0] = com.null_pad(sheet[xlsx.utils.encode_cell({ c: 25, r: irow })]);

        // 対応者コードを先頭0埋めフォーマット
        shain[1] = shain[0].toString().padStart(3, "0");

        // 連想配列より、キーから値を取得（社員コード⇒社員名）
        shain[2] = com.hash_syain[shain[1]];

        // ハッシュに値に該当するキーがない場合の対応       
        wdata[2] = (shain[2] == undefined) ? "" : shain[2];

        wdata[0] = com.null_pad(sheet[xlsx.utils.encode_cell({ c: 0, r: irow })]);                      // ID
        wdata[1] = wdata[0].toString().padStart(6, "0");
        wdata[3] = wdata[1] + "," + wdata[2];                                                       // 対応者（主）
        wdata[4] = wdata[3] + ",";
        wdata[5] = wdata[4] + com.null_pad(sheet[xlsx.utils.encode_cell({ c: 25, r: irow })]);          // 対応者コード

        fs.writeSync(fd, wdata[3] + "\n", 0);                                                       // ファイル書き込み
    }
    fs.closeSync(fd);
}




// ------------------------------------------------------------------
//
// ファイルの文字コードを調べる
// UTF-8であれば開く。UTF-8以外であれば、警告メッセージを出してエラーにする
//
// ------------------------------------------------------------------

function proc9_check_file(flg) {

    var fs = require('fs'), fd;
    var jschardet = require('jschardet');

    var filnm = ["sample_utf8", "sample_utf16", "sample_utf16be", "sample_sjis"];
    var file = __dirname + "/端末ID/" + filnm[flg] + ".txt";

    var text = fs.readFileSync(file);

    // 文字コードがUTF-8(ascii)以外であれば、エラーにする
    if (jschardet.detect(text).encoding != "ascii") {

        let msg1;
        msg1 = "ファイルの文字コードがUTF8でないため開くことができません。" + "\n";
        msg1 += "ファイルの文字コードは、" + jschardet.detect(text).encoding + "です。";
        console.log(msg1);

        let msg2;
        msg2 = jschardet.detect(text);
        console.log(msg2);
        return;
    }

    console.log("OK! ファイルは、" + jschardet.detect(text).encoding + "です。");
}



// ------------------------------------------------------------------
//
// 保守契約管理を、AppSuite→楽楽に移行するときの、「保守契約範囲」項目コンバート
// 保守契約範囲が複数選択されている場合、Enter区切→半角ブランク区切に置換する
//
// ------------------------------------------------------------------

function proc10_conv_hosyuhani() {

    var fs = require('fs');
    var fd;

    var file_red = __dirname + "/保守契約/保守契約管理アプリ_2回目" + ".csv";
    var file_wrt = __dirname + "/保守契約/保守契約管理アプリ_2回目_コンバート後" + ".csv";

    var line = [];

    // 読み込むファイルをオープンする
    var text = fs.readFileSync(file_red, 'utf8');

    line[0] = text.toString();

    // 対応と原因の間に改行コードがある場合、楽楽特有の何かのコードに置換
    line[1] = line[0].replace(/対応\r?\n原因/g, "対応	原因");
    line[2] = line[1].replace(/対応\r?\n現地/g, "対応	現地");

    // 文字列⇒配列
    line[3] = line[2].split('\n');

    // 書き込むファイルをオープンする
    fs.writeFileSync(file_wrt, "");
    fd = fs.openSync(file_wrt, "a");

    for (var idx in line[3]) {

        // １列単位で、文字列⇒配列
        let wdata = line[3][idx] + "";

        // 空白行は書き込まない
        if (wdata != "") {
            fs.writeSync(fd, wdata + "\n", 0);
        }
    }
    fs.closeSync(fd);
}





// ------------------------------------------------------------------
//
// 端末ID管理と請求処理の連携テストをするためのテストデータ（CSV）を作成する
// 引数:0 データをクリア
// 引数:1 データセット
//
// ------------------------------------------------------------------
function proc11_tanid_testdata_create(flg) {

    //var ary_1_data = [2, 5, 13, 14, 77, 82, 84, 85, 86, 87];     // 有償開始年月
    //var ary_2_data = [3, 5,  6, 14, 78, 82, 83, 85, 88];         // 有償終了年月
    //var ary_3_data = [4, 6, 13, 14, 80, 83, 84, 85, 87, 89];     // 斡旋手数料

    var ary_1_data = [2, 5, 13, 14];        // 有償開始年月
    var ary_2_data = [3, 5, 6, 14];         // 有償終了年月
    var ary_3_data = [4, 6, 13, 14];        // 斡旋手数料

    var fs = require('fs'), fd;

    // 書き込み先パス
    var file_path = __dirname + "/端末ID";

    // 書き込み先ファイル
    if (flg != 0) {
        var file_kais = "/端末ID_有償開始" + ".csv";
        var file_syuu = "/端末ID_有償終了" + ".csv";
        var file_asse = "/端末ID_斡旋手数" + ".csv";
    } else {
        var file_kais = "/端末ID_有償開始_クリア" + ".csv";
        var file_syuu = "/端末ID_有償終了_クリア" + ".csv";
        var file_asse = "/端末ID_斡旋手数_クリア" + ".csv";
    }

    var str_kais = new Date("2022/01/01");                // 有償開始の初期値
    var str_syuu = new Date("2023/06/01");                // 終了年月の初期値
    var str_asse = 1000;                                  // 斡旋手数料の初期値


    // ------------------------------------------------------------------------
    // 新規ファイルをオープン（有償開始年月）
    fs.writeFileSync(file_path + file_kais, "");
    var fd = fs.openSync(file_path + file_kais, "a");

    // ヘッダ部を出力
    fs.writeSync(fd, ["自動採番", "有償開始年月"].join(",") + "\n", 0);

    for (let idx in ary_1_data) {

        let ary_out = [];
        let set_date;

        if (flg != 0) {
            str_kais.setMonth(str_kais.getMonth() + 1);
            set_date = str_kais.getFullYear().toString() + "/" + (str_kais.getMonth() + 1).toString();
        } else {
            set_date = "\"\"";
        }

        ary_out.push(("00000" + ary_1_data[idx].toString()).slice(-5));
        ary_out.push(set_date);

        fs.writeSync(fd, ary_out.join(",") + "\n", 0);
    }
    fs.closeSync(fd);

    // ------------------------------------------------------------------------
    // 新規ファイルをオープン（有償終了年月）
    fs.writeFileSync(file_path + file_syuu, "");
    var fd = fs.openSync(file_path + file_syuu, "a");

    // ヘッダ部を出力    
    fs.writeSync(fd, ["自動採番", "有償終了年月"].join(",") + "\n", 0);

    for (let idx in ary_2_data) {

        let ary_out = [];
        let set_date;

        if (flg != 0) {
            str_syuu.setMonth(str_syuu.getMonth() + 1);
            set_date = str_syuu.getFullYear().toString() + "/" + (str_syuu.getMonth() + 1).toString();
        } else {
            set_date = "\"\"";
        }

        ary_out.push(("00000" + ary_2_data[idx].toString()).slice(-5));
        ary_out.push(set_date);

        fs.writeSync(fd, ary_out.join(",") + "\n", 0);
    }
    fs.closeSync(fd);

    // ------------------------------------------------------------------------
    // 新規ファイルをオープン（斡旋手数料）
    fs.writeFileSync(file_path + file_asse, "");
    var fd = fs.openSync(file_path + file_asse, "a");

    // ヘッダ部を出力
    fs.writeSync(fd, ["自動採番", "斡旋手数料"].join(",") + "\n", 0);

    for (let idx in ary_3_data) {

        let ary_out = [];
        let set_date;

        if (flg != 0) {
            set_date = str_asse + Number(idx) * 10;
        } else {
            set_date = "\"\"";
        }

        ary_out.push(("00000" + ary_3_data[idx].toString()).slice(-5));
        ary_out.push(set_date);

        fs.writeSync(fd, ary_out.join(",") + "\n", 0);
    }
    fs.closeSync(fd);
}





// ------------------------------------------------------------------
// 
// 端末ID管理の有償開始年月、有償終了年月のインポート用ファイルをクリアするための処理
//
// ------------------------------------------------------------------

function proc12_update_text() {

    var fs = require('fs'), fd;

    var path = __dirname + "/端末ID/";                               // パス
    var file = "請求インポート_有償開始終了年月_202206.csv";         // ファイル（元）
    var filn = file.replace('.', '_new.');                           // ファイル（先）

    // 読み込む
    var text = fs.readFileSync(path + file, 'utf8');                // 第二引数はテキストファイルの文字コードを指定
    var lines = text.toString().split('\r\n');                       // 改行コードで、文字列⇒1レコード:1要素の配列に変換

    // 空ファイルを新規作成
    fs.writeFileSync(path + filn, "");
    var fd = fs.openSync(path + filn, "a");

    for (var idx in lines) {

        let ary = [];

        if (idx == 0) {

            ary.push('"自動採番"');
            ary.push('"有償開始年月"');
            ary.push('"有償終了年月"');

        } else {

            let row = lines[idx].split(",");

            if (row[0] == "") { continue; }

            ary.push(row[0]);
            ary.push('""');
            ary.push('""');
        }
        // 書き込む
        fs.writeSync(fd, ary.join(",") + "\n", 0);
    }
    fs.closeSync(fd);
}


// ------------------------------------------------------------------
//
// 楽楽　見積テーブルのテスト用サンプルデータ作成
// 
// ------------------------------------------------------------------

function proc13_mitumori_table_testcsv_create() {

    var fs = require('fs');
    var fd;
    var file = __dirname + "/見積テーブル_テストデータ";

    fs.writeFileSync(file + "03.csv", "");                          // 空のファイルを書き
    fd = fs.openSync(file + "03.csv", "a");                         // ファイルをオープン

    // ファイルの存在チェック
    if (!fs.existsSync(file + "02.csv")) {
        console.log("読み込み元のCSVファイルが存在しません");
        return;
    }

    var text = fs.readFileSync(file + "02.csv", 'utf8');           // 第二引数はテキストファイルの文字コードを指定
    var lines = text.toString().split(/\r\n|\r|\n/);               // \r\n 文字列⇒配列

    for (let idx in lines) {

        let data = "";
        let arry = [];

        if (idx == 0) {

            // ヘッダ部にタイトルを出力
            arry.push("発行日");
            arry.push("捺印区分");
            arry.push("請求先");
            arry.push("件名（上段）");
            arry.push("件名（下段）");
            arry.push("納入場所");
            arry.push("納期");
            arry.push("お支払条件");
            arry.push("見積有効期限");
            arry.push("摘要（上段）");
            arry.push("摘要（下段）");
            arry.push("単位");
            arry.push("備考");

        } else {

            data = (idx < 4) ? "" : "2022年07月20日";
            arry.push(data);

            data = (idx < 4) ? "会社印" : "無";
            arry.push(data);

            arry.push("〇〇〇〇市　市長　米原　太郎");
            arry.push("クラウド型遠隔監視システム");
            arry.push("情報配信サービス利用契約");
            arry.push("△□×市　〇〇〇〇〇処理施設");
            arry.push("令和４年度内");
            arry.push("月末締め、翌月末支払い");
            arry.push("発行日から60日");
            arry.push("クラウド型遠隔監視システム");
            arry.push("情報配信サービス利用料");

            switch (true) {
                case ["1", "2", "3"].includes(idx.toString()):
                    data = "式";
                    break;
                case ["4", "5", "6"].includes(idx.toString()):
                    data = "ヶ月";
                    break;
                default:
                    data = "ヶ所";
            }
            arry.push(data);

            arry.push("備考の印字テストです。");
        }
        if (lines[idx] != "") { fs.writeSync(fd, lines[idx] + "," + arry.join(",") + "\r\n", 0); }
    }
    fs.closeSync(fd);

}





// ------------------------------------------------------------------
//
// Excelファイル処理（メイン）
// Excelファイルを読み込んで、特定のセルの値を読み込み、新規のテキストファイルに書き込む
//
// ------------------------------------------------------------------

function proc14_mitumori_excel_output_main(syori_flg) {

    var fs = require('fs');
    var path = "./見積書/" + syori_flg;
    var cnt = 0;

    // フォルダの存在確認
    if (!fs.existsSync(path)) {
        console.log("指定したフォルダが存在しません");
        return false;
    }

    // フォルダ内の全ファイル/フォルダを取得
    // filesは配列で返る  Array.isArray(files)はtrueで返る
    fs.readdir(path, (err, files) => {

        for (let i = 0; i < files.length; i++) {

            // 拡張子を取得
            // xlsx以外は読み飛ばす
            if (files[i].split('.').pop() != "xlsx") { continue; }

            // ファイル名の先頭に~$が付くファイルは、ゴミファイルのため読み飛ばす
            if (files[i].match(/^~\$/)) { continue; }

            // Excelファイルを読み込みCSVファイルに出力する
            proc14_mitumori_excel_output_proc01(syori_flg, cnt, files[i].split('.').shift());
            cnt++;
        }
    })
}



// ------------------------------------------------------------------
// Excelファイル処理（処理１）
// Excelファイルを読み込んで、特定のセルの値を読み込み、新規のテキストファイルに書き込む
// 
// ------------------------------------------------------------------
function proc14_mitumori_excel_output_proc01(syori_flg, syori_cnt, excel_file) {

    var fs = require('fs');
    var file = __dirname + "/" + syori_flg + ".csv";

    // １回目のみ、空ファイルを書き出す
    if (syori_cnt == 0) {
        fs.writeFileSync(file, "");
    }

    // 既存のファイルをオープン
    var fd = fs.openSync(file, "a");

    // Excelを読み込む
    var xlsx = require('xlsx');
    var workbook = xlsx.readFile("./見積書/" + syori_flg + "/" + excel_file + ".xlsx");

    // 全シートを配列で取得
    var sheet_ary = workbook.SheetNames;

    var flg = false;
    var sheet_index;

    // シート名に見積書/請求書/見積明細書/請求明細書を含むシートの存在チェック
    for (let i = 0; i < sheet_ary.length; i++) {
        if (sheet_ary[i].includes(syori_flg)) {
            flg = true;
            sheet_index = i;
            break;
        }
    }

    // シート名に見積書/請求書/見積明細書/請求明細書を含むシートがなければエラー
    if (flg == false) {
        console.log(sheet_ary);
        console.log("読み込み対象のシートが存在しません");
        return false;
    }

    // 該当のシート名を取得    
    var sheet_nm = workbook.SheetNames[sheet_index];

    // シートの全データをJSONで取得
    var sheet = workbook.Sheets[sheet_nm];

    var ary = [];
    let data = [];

    // タイトル部出力
    if (syori_cnt == 0) {

        switch (syori_flg) {
            case "見積書":

                ary.push("ファイル名");
                ary.push("自動採番");
                ary.push("自動採番（請求）");
                ary.push("請求先");
                ary.push("金額");
                ary.push("消費税");
                ary.push("合計");
                break;

            case "請求書":

                ary.push("ファイル名");
                ary.push("自動採番");
                ary.push("請求先");
                ary.push("金額");
                ary.push("消費税");
                ary.push("合計");
                break;

            case "見積明細書":
            case "請求明細書":

                ary.push("ファイル名");
                ary.push("自動採番（見積T）");
                ary.push("自動採番（請求T）");
                ary.push("請求先");
                ary.push("施設名");
                ary.push("〇月請求分");
                ary.push("単価");
                ary.push("月数");
                break;
        }
        fs.writeSync(fd, ary.join(',') + "\r\n", 0);
    }

    // 明細部出力
    ary = [];
    ary.push(excel_file);                                                         // ファイル名

    switch (syori_flg) {

        case "見積書":

            // c:列 r:行
            data[0] = com.null_pad(sheet[xlsx.utils.encode_cell({ c: 0, r: 1 })]);
            data[1] = (data[0].replace(/（見積№：  /, "")).replace(/  号）/, "");
            ary.push(data[1]);                                                        // 自動採番

            ary.push(com.null_pad(sheet[xlsx.utils.encode_cell({ c: 10, r: 10 })]));  // 自動採番（請求）
            ary.push(com.null_pad(sheet[xlsx.utils.encode_cell({ c: 0, r: 3 })]));    // 請求先

            ary.push(com.null_pad(sheet[xlsx.utils.encode_cell({ c: 7, r: 21 })]));   // 金額 
            ary.push(com.null_pad(sheet[xlsx.utils.encode_cell({ c: 2, r: 18 })]));   // 消費税
            ary.push(com.null_pad(sheet[xlsx.utils.encode_cell({ c: 2, r: 17 })]));   // 合計
            break;

        case "請求書":

            ary.push(com.null_pad(sheet["S39"]));   // 自動採番
            ary.push(com.null_pad(sheet["C3"]));    // 請求先
            ary.push(com.null_pad(sheet["K12"]));   // 金額 
            ary.push(com.null_pad(sheet["H9"]));    // 消費税
            ary.push(com.null_pad(sheet["F6"]));    // 合計
            break;

        case "見積明細書":

            break;

        case "請求明細書":

            break;
    }
    fs.writeSync(fd, ary.join(',') + "\r\n", 0);
    fs.closeSync(fd);
}





// proc1_write_header();
// proc2_get_user();
// proc3_match_user();
// proc4_match_shet();
// proc5_write_sisetu();
// proc6_write_sisetu_only();
// proc7_write_zeropad();
// proc9_check_file(3);
// proc10_conv_hosyuhani();
// proc11_tanid_testdata_create();
// proc12_update_text();
// proc13_mitumori_table_testcsv_create();
// proc14_mitumori_excel_output_main("見積書");


