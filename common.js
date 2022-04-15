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


