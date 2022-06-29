var hash = {
    "0001": "〇",
    "0002": "□",
    "0003": "△"
};

// モジュール化して、外部ファイルから読み出せるようにする
exports.hash = hash;


// ------------------------------------------------------------------
// 社員コードと社員名のハッシュ
// ------------------------------------------------------------------

var hash_syain = {
    "032": "〇〇 太郎",
    "038": "△△ 花子" 
};


// モジュール化して、外部ファイルから読み出せるようにする
exports.hash_syain = hash_syain;


// ------------------------------------------------------------------
//
// undefinedの場合、""に変換
// Excelのセルが空白の場合、undefinedになる
// undefinedに対して、.vするとエラーになるためこのような対応が必要
// proc5_write_sisetuから読み込まれるサブ関数
//
// ------------------------------------------------------------------

exports.null_pad = function (data) {

    if (data == undefined) {
        return "";
    } else {
        return data.v;
    }
}






// ------------------------------------------------------------------
//
// 日付がセットされるセルが、日付型 or 文字列型を判断する
// 日付型であれば、.vするとシリアル値が取得される。.wすると、dd/mm/yyのフォーマットで日付が取得されるため変換する
// 文字列であれば、.vで素直に取得できる
// proc5_write_sisetuから読み込まれるサブ関数
//
// ------------------------------------------------------------------

exports.get_date = function (data) {

    var wdata;

    if (data == undefined) {

        wdata = "";

    } else {
        switch (typeof data.v) {

            // Numberが返ってくれば、シリアル値
            case "number":

                let data_d = new Date(data.w);
                wdata = data_d.getFullYear() + "/" + Number(data_d.getMonth() + 1) + "/" + data_d.getDate();
                break;

            // Stringが返ってくれば、文字列
            case "string":

                wdata = data.v;
                break;

            default:
                wdata = "想定外のデータ型がキャッチされた。セルにどんなデータがセットされているか調べること";
                console.log(wdata);
        }
    }
    return wdata;
}



// ------------------------------------------------------------------
// 
// ファイルの文字コードを判断して、S-JISであれば、UTF-8に変換した
// ファイルを、_utfをつけて新規作成する。戻り値はUTF-8であれば0、S-JIS
// であれば1、それ以外の文字コードであれば、-1を返す
//
// （使い方例）
// var path = __dirname + "/端末ID/";                         // パス
// var file = "請求インポート_有償開始終了年月_202206.csv";   // ファイル（変換元）
// var ret  = test_check_utf_file(path + file);
//
// ------------------------------------------------------------------

exports.check_utf_file = function (file) {

    var fs = require('fs'), fd;
    var jschardet = require('jschardet');                         // 文字コードを判断
    var iconv = require('iconv-lite');                            // 文字コードを変換
    var msg;

    var filn = file.replace('.', '_utf.');
    var text = fs.readFileSync(file);                             //元ファイルの内容を取得

    // 文字コードがUTF-8(ascii)以外であれば、エラーにする
    switch (jschardet.detect(text).encoding) {

        case "ascii":                                             // UTF-8(ascii)

            return 0;

        case "UTF-8":

            return 0;

        case "SHIFT_JIS":

            // 空ファイルを作成
            fs.writeFileSync(filn, "");
            var fd = fs.openSync(filn, "a");

            // 文字コード変換 ms932 -> utf8
            var data = iconv.decode(text, 'ms932');

            // 新規ファイルを書き込む
            fs.writeSync(fd, data, 0)
            fs.closeSync(fd);

            msg = "ファイルはS-JISです。UTF-8に変換したファイルを作成しました"
            console.log(msg);
            return 1;

        default:
            msg = "文字コードは" + jschardet.detect(text).encoding + "です。S-JISでもUTF-8でもない文字コードです"
            console.log(msg);
            return -1;
    }
}

